###
# NV Speech Player - NVDA synth driver (modernized)
#
# What this patch changes:
# - Uses explicit ctypes prototypes (via speechPlayer.py) and the ms->samples conversion the DLL expects.
# - Adds a background worker so speak() doesn't block NVDA while doing phoneme conversion.
# - Uses a WavePlayer.feed call style like sv.py, with a fallback for older NVDA.
# - Schedules synthDoneSpeaking after playback ends (using a 0-byte feed callback).
###

import os
import re
import threading
import queue
import math
import ctypes
import weakref
from collections import OrderedDict

import config
import nvwave

from logHandler import log
from synthDrivers import _espeak
from synthDriverHandler import SynthDriver, VoiceInfo, synthIndexReached, synthDoneSpeaking

# NVDA command classes moved around across versions; keep imports tolerant.
try:
	from speech.commands import IndexCommand, PitchCommand
except Exception:
	try:
		from speech.commands import IndexCommand  # type: ignore
		PitchCommand = None  # type: ignore
	except Exception:
		import speech  # fallback
		IndexCommand = getattr(speech, "IndexCommand", None)
		PitchCommand = getattr(speech, "PitchCommand", None)

from autoSettingsUtils.driverSetting import NumericDriverSetting

from . import speechPlayer
from . import ipa


re_textPause = re.compile(r"(?<=[.?!,:;])\s", re.DOTALL | re.UNICODE)


# Voice presets (unchanged idea, just cleaned a bit)
voices = {
	"Adam": {
		"cb1_mul": 1.3,
		"pa6_mul": 1.3,
		"fricationAmplitude_mul": 0.85,
	},
	"Benjamin": {
		"cf1_mul": 1.01,
		"cf2_mul": 1.02,
		"cf4": 3770,
		"cf5": 4100,
		"cf6": 5000,
		"cfNP_mul": 0.9,
		"cb1_mul": 1.3,
		"fricationAmplitude_mul": 0.7,
		"pa6_mul": 1.3,
	},
	# Keep BOTH spellings so existing configs don’t break.
	"Caleb ": {
		"aspirationAmplitude": 1,
		"voiceAmplitude": 0,
	},
	"Caleb": {
		"aspirationAmplitude": 1,
		"voiceAmplitude": 0,
	},
	"David": {
		"voicePitch_mul": 0.75,
		"endVoicePitch_mul": 0.75,
		"cf1_mul": 0.75,
		"cf2_mul": 0.85,
		"cf3_mul": 0.85,
	},
}


def applyVoiceToFrame(frame, voiceName):
	v = voices.get(voiceName) or voices.get("Adam", {})
	for paramName in (x[0] for x in frame._fields_):
		absVal = v.get(paramName)
		if absVal is not None:
			setattr(frame, paramName, absVal)
		mulVal = v.get("%s_mul" % paramName)
		if mulVal is not None:
			setattr(frame, paramName, getattr(frame, paramName) * mulVal)


class _BgThread(threading.Thread):
	def __init__(self, q, stopEvent):
		super().__init__(name=f"{self.__class__.__module__}.{self.__class__.__qualname__}")
		self.daemon = True
		self._q = q
		self._stop = stopEvent

	def run(self):
		while not self._stop.is_set():
			try:
				item = self._q.get(timeout=0.2)
			except queue.Empty:
				continue
			try:
				if item is None:
					return
				func, args, kwargs = item
				func(*args, **kwargs)
			except Exception:
				log.error("nvSpeechPlayer: error in background thread", exc_info=True)
			finally:
				try:
					self._q.task_done()
				except Exception:
					pass


class _AudioThread(threading.Thread):
	def __init__(self, synth, player, sampleRate):
		super().__init__(name=f"{self.__class__.__module__}.{self.__class__.__qualname__}")
		self.daemon = True
		self._synthRef = weakref.ref(synth)
		self._player = player
		self._sampleRate = int(sampleRate)

		self._keepAlive = True
		self.isSpeaking = False

		self._wake = threading.Event()
		self._init = threading.Event()

		self._wavePlayer = None
		self._outputDevice = None

		self.start()
		self._init.wait()

	def _getOutputDevice(self):
		try:
			return config.conf["speech"]["outputDevice"]
		except Exception:
			try:
				return config.conf["audio"]["outputDevice"]
			except Exception:
				return None

	def _feed(self, data, onDone=None):
		"""Modern NVDA: feed(data, length, onDone=...)
		Older NVDA: feed(data, onDone=...)
		"""
		if not self._wavePlayer:
			return
		try:
			self._wavePlayer.feed(data, len(data), onDone=onDone)
		except TypeError:
			try:
				if onDone is None:
					self._wavePlayer.feed(data)
				else:
					self._wavePlayer.feed(data, onDone=onDone)
			except Exception:
				pass

	def terminate(self):
		self._keepAlive = False
		self.isSpeaking = False
		self._wake.set()
		self.join(timeout=2.0)
		try:
			if self._wavePlayer:
				self._wavePlayer.stop()
		except Exception:
			pass

	def kick(self):
		self._wake.set()

	def run(self):
		try:
			self._outputDevice = self._getOutputDevice()
			self._wavePlayer = nvwave.WavePlayer(
				channels=1,
				samplesPerSec=self._sampleRate,
				bitsPerSample=16,
				outputDevice=self._outputDevice,
			)
		finally:
			self._init.set()

		lastNotifiedIndex = None

		while self._keepAlive:
			self._wake.wait()
			self._wake.clear()

			while self._keepAlive and self.isSpeaking:
				data = self._player.synthesize(8192)
				if data:
					n = int(getattr(data, "length", 0) or 0)
					if n <= 0:
						continue
					nbytes = n * ctypes.sizeof(ctypes.c_short)
					audioBytes = ctypes.string_at(ctypes.addressof(data), nbytes)

					idx = int(self._player.getLastIndex())
					if idx >= 0 and idx != lastNotifiedIndex:
						s = self._synthRef()
						def cb(index=idx, synth=s):
							if synth:
								synthIndexReached.notify(synth=synth, index=index)
						self._feed(audioBytes, onDone=cb)
						lastNotifiedIndex = idx
					else:
						self._feed(audioBytes)
					continue

				break

			# Done: notify only after playback drains.
			if self._keepAlive and self._wavePlayer:
				s = self._synthRef()
				if s:
					def doneCb(synth=s):
						synthDoneSpeaking.notify(synth=synth)
					self._feed(b"", onDone=doneCb)
				try:
					self._wavePlayer.idle()
				except Exception:
					pass

			self.isSpeaking = False


class SynthDriver(SynthDriver):
	name = "nvSpeechPlayer"
	description = "NV Speech Player"

	supportedSettings = (
		SynthDriver.VoiceSetting(),
		SynthDriver.RateSetting(),
		SynthDriver.PitchSetting(),
		SynthDriver.VolumeSetting(),
		SynthDriver.InflectionSetting(),
	)

	supportedCommands = {c for c in (IndexCommand, PitchCommand) if c}
	supportedNotifications = {synthIndexReached, synthDoneSpeaking}

	exposeExtraParams = True
	_ESPEAK_PHONEME_MODE = 0x36100 + 0x82

	def __init__(self):
		super().__init__()

		# If you're running 64-bit NVDA, a 32-bit speechPlayer.dll will not work.
		if ctypes.sizeof(ctypes.c_void_p) != 4:
			raise RuntimeError("nvSpeechPlayer: 32-bit only (speechPlayer.dll is x86)")

		if self.exposeExtraParams:
			self._extraParamNames = [x[0] for x in speechPlayer.Frame._fields_]
			extraSettings = tuple(
				NumericDriverSetting(f"speechPlayer_{x}", f"Frame: {x}")
				for x in self._extraParamNames
			)
			self.supportedSettings = self.supportedSettings + extraSettings
			for x in self._extraParamNames:
				setattr(self, f"speechPlayer_{x}", 50)

		self._sampleRate = 16000
		self._player = speechPlayer.SpeechPlayer(self._sampleRate)

		_espeak.initialize()
		_espeak.setVoiceByLanguage("en")

		self._curPitch = 50
		self._curVoice = "Adam"
		self._curInflection = 0.5
		self._curVolume = 1.0
		self._curRate = 1.0

		self._audio = _AudioThread(self, self._player, self._sampleRate)

		self._bgQueue = queue.Queue()
		self._bgStop = threading.Event()
		self._bgThread = _BgThread(self._bgQueue, self._bgStop)
		self._bgThread.start()

		# apply defaults
		self.pitch = 50
		self.rate = 50
		self.volume = 90
		self.inflection = 60

	@classmethod
	def check(cls):
		if ctypes.sizeof(ctypes.c_void_p) != 4:
			return False
		dllPath = os.path.join(os.path.dirname(__file__), "speechPlayer.dll")
		return os.path.isfile(dllPath)

	def _enqueue(self, func, *args, **kwargs):
		if self._bgStop.is_set():
			return
		self._bgQueue.put((func, args, kwargs))

	def _notifyIndexesAndDone(self, indexes):
		for i in indexes:
			synthIndexReached.notify(synth=self, index=i)
		synthDoneSpeaking.notify(synth=self)

	def _espeakTextToIPA(self, text):
		if not text:
			return ""
		textBuf = ctypes.create_unicode_buffer(text)
		textPtr = ctypes.c_void_p(ctypes.addressof(textBuf))

		chunks = []
		lastPtr = None
		while textPtr and textPtr.value:
			if lastPtr == textPtr.value:
				break
			lastPtr = textPtr.value
			phonemeBuf = _espeak.espeakDLL.espeak_TextToPhonemes(
				ctypes.byref(textPtr),
				_espeak.espeakCHARS_WCHAR,
				self._ESPEAK_PHONEME_MODE,
			)
			if phonemeBuf:
				chunks.append(ctypes.string_at(phonemeBuf))
			else:
				break

		ipaBytes = b"".join(chunks)
		try:
			ipaText = ipaBytes.decode("utf8", errors="ignore")
		except Exception:
			ipaText = ""

		ipaText = ipaText.strip()
		ipaText = ipaText.replace("ə͡l", "ʊ͡l")
		ipaText = ipaText.replace("a͡ɪ", "ɑ͡ɪ")
		ipaText = ipaText.replace("e͡ɪ", "e͡i")
		ipaText = ipaText.replace("ə͡ʊ", "o͡u")
		return ipaText.strip()

	def speak(self, speechSequence):
		indexes = []
		anyText = False
		for item in speechSequence:
			if IndexCommand and isinstance(item, IndexCommand):
				indexes.append(item.index)
			elif isinstance(item, str) and item.strip():
				anyText = True

		if (not anyText) and indexes:
			self._enqueue(self._notifyIndexesAndDone, indexes)
			return
		if not anyText and not indexes:
			self._enqueue(self._notifyIndexesAndDone, [])
			return

		self._enqueue(self._speakBg, list(speechSequence))

	def _speakBg(self, speakList):
		userIndex = None
		pitchOffset = 0

		# Merge adjacent strings
		i = 0
		while i < len(speakList):
			item = speakList[i]
			if i > 0:
				prev = speakList[i - 1]
				if isinstance(item, str) and isinstance(prev, str):
					speakList[i - 1] = " ".join([prev, item])
					del speakList[i]
					continue
			i += 1

		endPause = 20.0

		for item in speakList:
			if PitchCommand and isinstance(item, PitchCommand):
				pitchOffset = getattr(item, "offset", 0) or 0
				continue
			if IndexCommand and isinstance(item, IndexCommand):
				userIndex = item.index
				continue
			if not isinstance(item, str):
				continue

			for chunk in re_textPause.split(item):
				if not chunk:
					continue
				chunk = chunk.strip()
				if not chunk:
					continue

				clauseType = chunk[-1]
				if clauseType in (".", "!"):
					endPause = 150.0
				elif clauseType == "?":
					endPause = 150.0
				elif clauseType == ",":
					endPause = 120.0
				else:
					endPause = 100.0
					clauseType = None

				endPause = float(endPause) / float(self._curRate)

				ipaText = self._espeakTextToIPA(chunk)
				if not ipaText:
					continue

				pitch = float(self._curPitch) + float(pitchOffset)
				basePitch = 25.0 + (21.25 * (pitch / 12.5))

				for frame, frameDuration, fadeDuration in ipa.generateFramesAndTiming(
					ipaText,
					speed=self._curRate,
					basePitch=basePitch,
					inflection=self._curInflection,
					clauseType=clauseType,
				):
					if frame:
						applyVoiceToFrame(frame, self._curVoice)

						if self.exposeExtraParams:
							for x in self._extraParamNames:
								ratio = float(getattr(self, f"speechPlayer_{x}", 50)) / 50.0
								setattr(frame, x, getattr(frame, x) * ratio)

						frame.preFormantGain *= self._curVolume

					self._player.queueFrame(frame, frameDuration, fadeDuration, userIndex=userIndex)
					userIndex = None

		self._player.queueFrame(None, endPause, max(10.0, 10.0 / float(self._curRate)), userIndex=userIndex)

		self._audio.isSpeaking = True
		self._audio.kick()

	def cancel(self):
		try:
			self._player.queueFrame(None, 20.0, 5.0, purgeQueue=True)
		except Exception:
			pass
		try:
			self._audio.isSpeaking = False
			self._audio.kick()
		except Exception:
			pass
		try:
			if self._audio and self._audio._wavePlayer:
				self._audio._wavePlayer.stop()
		except Exception:
			pass

	def pause(self, switch):
		try:
			if self._audio and self._audio._wavePlayer:
				self._audio._wavePlayer.pause(switch)
		except Exception:
			pass

	def terminate(self):
		try:
			self.cancel()
		except Exception:
			pass

		try:
			self._bgStop.set()
			try:
				self._bgQueue.put(None)
			except Exception:
				pass
			try:
				self._bgThread.join(timeout=2.0)
			except Exception:
				pass
		except Exception:
			pass

		try:
			self._audio.terminate()
		except Exception:
			pass

		try:
			self._player.terminate()
		except Exception:
			pass

		try:
			_espeak.terminate()
		except Exception:
			pass

	def _get_rate(self):
		return int(math.log(self._curRate / 0.25, 2) * 25.0)

	def _set_rate(self, val):
		self._curRate = 0.25 * (2 ** (float(val) / 25.0))

	def _get_pitch(self):
		return int(self._curPitch)

	def _set_pitch(self, val):
		self._curPitch = int(val)

	def _get_volume(self):
		return int(self._curVolume * 75)

	def _set_volume(self, val):
		self._curVolume = float(val) / 75.0

	def _get_inflection(self):
		return int(self._curInflection / 0.01)

	def _set_inflection(self, val):
		self._curInflection = float(val) * 0.01

	def _get_voice(self):
		return self._curVoice

	def _set_voice(self, voice):
		if voice not in self.availableVoices:
			voice = "Adam"
		self._curVoice = voice
		if self.exposeExtraParams:
			for paramName in self._extraParamNames:
				setattr(self, f"speechPlayer_{paramName}", 50)

	def _getAvailableVoices(self):
		d = OrderedDict()
		for name in sorted(voices):
			d[name] = VoiceInfo(name, name)
		return d
