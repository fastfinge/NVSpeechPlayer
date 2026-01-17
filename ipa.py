###
#This file is a part of the NV Speech Player project. 
#URL: https://bitbucket.org/nvaccess/speechplayer
#Copyright 2014 NV Access Limited.
#This program is free software: you can redistribute it and/or modify
#it under the terms of the GNU General Public License version 2.0, as published by
#the Free Software Foundation.
#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
#This license can be found at:
#http://www.gnu.org/licenses/old-licenses/gpl-2.0.html
###

import os
import itertools
import codecs
import ast
import re
try:
	from . import speechPlayer
except Exception:
	import speechPlayer

_here=os.path.dirname(__file__)
dataPath=os.path.join(_here,'data.py')
_altDataPath=os.path.join(_here,'data_full.py')
if os.path.exists(_altDataPath):
	dataPath=_altDataPath

data=ast.literal_eval(codecs.open(dataPath,'r','utf8').read())

# --- Prosody / cadence tuning ---
#
# This build restores the older (ipa-older.py) timing/fade model by default
# across languages, while keeping the newer multi-language IPA normalization
# and phoneme mapping rules intact.

#
# The multi-language version of this file reduced how much stressed syllables
# are slowed down (primary stress: /1.25, secondary: /1.07). The older
# NVSpeechPlayer-era tuning (primary: /1.4, secondary: /1.1) tends to produce
# a more 'shaped' cadence and clearer syllable peaks (often described as more
# ETI Eloquence-like).
#
# To keep newer language/IPA normalization improvements while restoring the
# older cadence, we select stress slowdowns by language, defaulting to the
# newer factors for non-English languages. You can extend this table if you
# want the older cadence for other languages as well.

_STRESS_SLOWDOWN_BY_LANG = {
	'default': (1.4, 1.1),
	# Older cadence for English variants.
	'en': (1.4, 1.1),
	'en-us': (1.4, 1.1),
	'en-gb': (1.4, 1.1),
	'en-uk': (1.4, 1.1),
	'en-ca': (1.4, 1.1),
	'en-us-nyc': (1.4, 1.1),

	# Newer/milder stress shaping for non-English languages.
	# This keeps the 'English cadence' from bleeding into other voices.
	'hu': (1.25, 1.07),
	'pl': (1.25, 1.07),
	'es': (1.25, 1.07),
	'pt': (1.25, 1.07),
	'fr': (1.25, 1.07),
	'de': (1.25, 1.07),
	'it': (1.25, 1.07),
	'da': (1.25, 1.07),
	'ro': (1.25, 1.07),
}

def _get_stress_slowdown(language=None):
	"""Return (primaryDiv, secondaryDiv) for stress-based speed scaling.

	Language matching prefers the most specific tag first, then falls back to
	the base language, then to 'default'.	
	"""
	lang = (language or '').replace('_', '-').lower()
	if not lang:
		return _STRESS_SLOWDOWN_BY_LANG['default']
	parts = lang.split('-')
	for i in range(len(parts), 0, -1):
		key = '-'.join(parts[:i])
		val = _STRESS_SLOWDOWN_BY_LANG.get(key)
		if val:
			return val
	return _STRESS_SLOWDOWN_BY_LANG.get(parts[0], _STRESS_SLOWDOWN_BY_LANG['default'])

def normalizeIPA(text, language=None):
	"""Normalize eSpeak phoneme/IPA output into a stable IPA stream.

	Accepts either eSpeak phoneme mnemonics (-x) or true IPA output (--ipa),
	then maps/normalizes into symbols that exist in data.py.
	"""
	if text is None:
		return u""
	if not isinstance(text, str):
		# Best effort decoding.
		try:
			text = text.decode("utf-8", "ignore")
		except Exception:
			text = str(text)

	lang = (language or "").replace("_", "-").lower()
	isEnglish = lang.startswith("en")

	# eSpeak's default English voice "en" is British (non-rhotic).
	# Keep rhotic/non-rhotic decisions purely based on the language tag so we can
	# make "world" and "broadcast" differ between US and UK even when eSpeak emits
	# the same mnemonic phonemes.
	#
	# Rhotic: en-us, en-ca, en-us-nyc, etc.
	# Non-rhotic: en, en-gb, en-uk.
	isNonRhoticEnglish = isEnglish and (lang == "en" or lang.startswith("en-gb") or lang.startswith("en-uk"))
	isRhoticEnglish = isEnglish and not isNonRhoticEnglish
	isBritishEnglish = isNonRhoticEnglish
	isHungarian = lang.startswith("hu")
	isPolish = lang.startswith("pl")
	isSpanish = lang.startswith("es")
	isPortuguese = lang.startswith("pt")
	isFrench = lang.startswith("fr")
	isGerman = lang.startswith("de")
	isItalian = lang.startswith("it")
	isDanish = lang.startswith("da")
	isRomanian = lang.startswith("ro")

	# --- eSpeak utility codes / cleanup ---
	# Normalise tie bar variants.
	text = text.replace(u"͜", u"͡")
	# Remove common wrapper punctuation.
	for c in (u"[", u"]", u"(", u")", u"{", u"}", u"/", u"\\"):
		text = text.replace(c, u"")
	# eSpeak dictionary utility codes.
	text = text.replace(u"||", u" ")
	text = text.replace(u"|", u"")
	text = text.replace(u"%", u"")
	text = text.replace(u"=", u"")
	text = text.replace(u"!", u"")
	# Pause / separator markers.
	text = text.replace(u"_:", u" ")
	text = text.replace(u"_", u" ")
	text = text.replace(u"-", u"")

	# Stress/length markers.
	text = text.strip().replace(u"'", u"ˈ").replace(u",", u"ˌ")
	text = text.replace(u":", u"ː")

	# Portuguese nasalization:
	# eSpeak uses '~' in -x mnemonics (e.g. &U~, o~, u~) and a combining tilde in --ipa.
	# Convert these into stable single-codepoint vowels that exist in data.py so they don't get stripped/dropped.
	if isPortuguese:
		# -x mnemonics (longest-first so '&U~' is handled before '&~').
		text = text.replace(u"&U~", u"ãᴜ")
		text = text.replace(u"U~", u"ᴜ")
		text = text.replace(u"&~", u"ã")
		text = text.replace(u"a~", u"ã")
		text = text.replace(u"o~", u"õ")
		text = text.replace(u"u~", u"ũ")
		text = text.replace(u"e~", u"ẽ")
		text = text.replace(u"i~", u"ĩ")

		# --ipa output (combining tilde).
		text = text.replace(u"ɐ̃ʊ̃", u"ãᴜ")
		text = text.replace(u"ɐ̃", u"ã")
		text = text.replace(u"õ", u"õ")
		text = text.replace(u"ũ", u"ũ")
		text = text.replace(u"ʊ̃", u"ᴜ")
		text = text.replace(u"ẽ", u"ẽ")
		text = text.replace(u"ĩ", u"ĩ")

	# Drop common IPA diacritics we don't model.
	# (palatalization, nasalization, etc.)
	text = text.replace(u"ʲ", u"")
	text = text.replace(u"̃", u"")
	text = text.replace(u"~", u"")

	# --- Multi-character replacements (longest-first) ---
	multi = {
		# eSpeak tap markers used by several languages.
		# "**" (and sometimes "*") indicates an alveolar tap.
		u"**": u"ɾ",
		u"*": u"ɾ",
		# Affricates (IPA sequences -> tied form so we hit the _isAfricate entries).
		u"tʃ": u"t͡ʃ",
		u"dʒ": u"d͡ʒ",
		u"tɕ": u"t͡ɕ",
		u"dʑ": u"d͡ʑ",
		# Mnemonic affricates.
		u"t͡S": u"t͡ʃ",
		u"d͡Z": u"d͡ʒ",
		u"ts": u"t͡s",
		u"dz": u"d͡z",
		# Polish palatalized affricates/fricatives (mnemonics).
		u"S;": u"ɕ",
		u"Z;": u"ʑ",
		u"ts;": u"t͡ɕ",
		u"dz;": u"d͡ʑ",
		# Spanish/Portuguese palatals.
		u"n^": u"ɲ",
		u"l^": u"ʎ",
		# Portuguese 'lh' often appears as lj.
		u"lj": u"ʎ" if isPortuguese else u"lj",
		# Rolled r markers.
		u"RR2": u"r",
		u"R2": u"r",
		# Unstressed/reduced English vowels.
		u"I2": u"ɪ",
		u"I#": (u"ᵻ" if isEnglish and isRhoticEnglish else u"ɪ"),
		u"I2#": (u"ᵻ" if isEnglish and isRhoticEnglish else u"ɪ"),
		u"e#": u"ɛ",
		# Syllabic /l/.
		u"@L": u"əl",
	}

	# German: ich-Laut appears as "C" in eSpeak mnemonics (e.g. 'ich', 'Mädchen').
	# Map it to ç if available; otherwise fall back to x so it won't drop.
	if isGerman:
		multi[u"C"] = (u"ç" if u"ç" in data else u"x")


	# Portuguese: strengthen common diphthongs and handle strong-R fallbacks.
	# eSpeak's -x mnemonics often use sequences like aI, eI, aU, eU, and ow.
	# Turning these into tied vowel pairs makes the glide audible without adding
	# an extra syllable.
	if isPortuguese:
		# Some eSpeak voices output "rr" for a strong R; map it to x so it won't drop.
		multi[u"rr"] = u"x"
		# Common Portuguese diphthongs.
		multi[u"aI"] = u"a͡i"
		multi[u"eI"] = u"e͡i"
		multi[u"oI"] = u"o͡i"
		multi[u"aU"] = u"a͡u"
		multi[u"eU"] = u"e͡u"
		multi[u"EU"] = u"ɛ͡u"
		# "ou" is typically emitted as "ow" in pt-br/pt-pt.
		multi[u"ow"] = u"o͡u"


	# Hungarian nuance: long é is a different vowel quality from short e.
	# eSpeak outputs long é as "e:" (lowercase e + length marker), which we
	# normalize earlier into "eː".
	# Map it to an internal vowel symbol so we can tune Hungarian é without
	# changing English/Spanish/Portuguese "e".
	if isHungarian:
		multi[u"eː"] = u"ᴇː"

	# Hungarian vowel length nuance:
	#
	# eSpeak Hungarian uses lowercase "a:" for long "á" (/aː). The short "a" phoneme is
	# the mnemonic "A" (/ɒ/).
	#
	# Some eSpeak outputs may contain "A:" ("Aː" after our ':'->'ː' rewrite) due to
	# phrase/word lengthening (for example the standalone article "a").
	#
	# Hungarian does not have a *phonemic* long /ɒ/ counterpart (the long "á" is /aː/ and
	# eSpeak normally emits it as lowercase "a:"). If we treat "A:" as a true long vowel,
	# the short "a" can end up sounding too close to "á" in some contexts.
	#
	# So, for Hungarian, collapse "Aː" back to "A". This keeps the rounded/back quality,
	# and avoids accidental lengthening of the short vowel.
	if isHungarian:
		multi[u"Aː"] = u"A"

	# Mnemonic tS/dZ affricates. eSpeak uses these as /tʃ/ and /dʒ/ across voices.
	multi[u"tS"] = u"t͡ʃ"
	multi[u"dZ"] = u"d͡ʒ"

	# English: tie common diphthongs and handle flap markers.
	# This prevents over-long vowels and lets us differentiate US/CA vs UK-style
	# diphthongs where eSpeak uses the same mnemonic sequences.
	if isEnglish:
		# Common English diphthongs in eSpeak mnemonics.
		# Use 'ɑ' starts for PRICE/MOUTH so later TRAP mapping doesn't affect them.
		multi[u"aI"] = u"ɑ͡ɪ"
		multi[u"aU"] = u"ɑ͡ʊ"
		multi[u"OI"] = u"ɔ͡ɪ"
		# eSpeak uses "aa" for the BATH/PALM set across English variants.
		# Dialect differs:
		# - Rhotic (US/CA): closer to TRAP /æ/ (bath, cast).
		# - Non-rhotic (UK): broad /ɑː/ (bath, cast).
		multi[u"aa"] = (u"æ" if isRhoticEnglish else u"ɑː")
		# GOAT: US/CA tends to 'oʊ', UK and many non‑rhotic accents closer to 'əʊ'.
		multi[u"oU"] = (u"o͡ʊ" if isRhoticEnglish else u"ə͡ʊ")
		# FACE: give US/CA a slightly tenser offglide (towards i) for a clearer diphthong.
		multi[u"eI"] = (u"e͡i" if isRhoticEnglish else u"e͡ɪ")
		# eSpeak mnemonic flap markers only for rhotic (US/CA) accents.
		if isRhoticEnglish:
			multi[u"t#"] = u"ɾ"
			multi[u"d#"] = u"ɾ"

	# eSpeak --ipa output for English commonly uses ɜ(ː) for stressed NURSE even in rhotic accents.
	# Convert it to the r-coloured vowel ɝ(ː) for rhotic English so words like 'world' don't sound UK-ish.
	if isEnglish and isRhoticEnglish:
		multi[u"ɜː"] = u"ɝː"
		multi[u"ɜ"] = u"ɝ"

	# English rhotic clusters.

# English rhotic clusters and vowel normalization
	if isEnglish:
		if isRhoticEnglish:
			multi.update({
				u"3ː": u"ɝː",   # US NURSE (thirty)
				u"3": u"ɚ",    # US unstressed rhotic (better)
				u"A@": u"ɑɹ",  # US START
				u"O@": u"ɔːɹ",  # US NORTH
				u"o@": u"ɔːɹ",
				u"i@3": u"ɪɹ",
				u"i@": u"ɪɹ",
				u"e@": u"ɛɹ",  # US SQUARE
			})
		else:
			multi.update({
				u"3ː": u"ɜː",   # UK NURSE (long vowel, no R)
				u"3": u"ə",    # UK unstressed (schwa)
				u"A@": u"ɑː",
				# Keep a distinct UK THOUGHT/NORTH vowel when data provides it.
				# eSpeak mnemonics commonly use O@/o@ for "or"/"ore" sets (north, four).
				# If we always map this to "ɔ", UK and US can collapse into the same sound.
				u"O@": u"Oː" if (u"O" in data) else u"ɔː",
				u"o@": u"Oː" if (u"O" in data) else u"ɔː",
				u"i@3": u"ɪə",
				u"i@": u"ɪə",
				u"e@": u"ɛə",
			})
	if isGerman:
		# Map vocalic R symbols to our new data entry
		multi.update({
			u"ɐ": u"ɐ",
			u"ɐ̯": u"ɐ", # Handle the non-syllabic version
			u"R2": u"ɐ", # eSpeak mnemonic for vocalic R
			u"@2": u"ɐ", # Another common eSpeak variant
		})
	if isGerman:
		text = text.replace(u"ɐ̯", u"ɐ") # Strip the "non-syllabic" tail
	# Portuguese: some eSpeak variants may emit 'R' at the start of a word for the
	# strong /R/ (often realised as [h]/[x] in Brazilian Portuguese).
	# Only rewrite it at true word starts so clusters like 'bR' stay a tap/flap.
	if isPortuguese:
		text = re.sub(r'(^|\s)R', r'\1x', text)

	for k in sorted(multi, key=len, reverse=True):
		text = text.replace(k, multi[k])

	# eSpeak sometimes appends numeric allophone markers (for example, t2)
	# that are not phonemes in our table. After multi-replacements, drop
	# any remaining standalone '2' markers so they don't become unknown
	# phonemes (e.g. desktop -> d'Eskt20p).
	text = text.replace(u"2", u"")

	# --- Single-character ASCII mnemonics ---
	asciiMap = {
		u"@": u"ə",
		u"E": u"ɛ",
		# "O" is overloaded in eSpeak mnemonics.
		# - In English, we keep UK (non‑rhotic) "O" as its own vowel (rounded THOUGHT)
		#   when the table provides an entry for "O".
		# - In US/CA (rhotic) English, map it to "ɔ" (more open) for a more American feel.
		# - In Portuguese, some builds use "O" for ó/open-O; route to a dedicated phoneme if present.
		u"O": (
			u"ᴐ" if (isPortuguese and (u"ᴐ" in data))
			else (u"O" if (isEnglish and isNonRhoticEnglish and (u"O" in data)) else u"ɔ")
		),
		u"V": u"ʌ",
		u"U": (u"u" if isPortuguese else u"ʊ"),
		u"I": (u"i" if isPortuguese else u"ɪ"),
		u"J": u"j",
		# eSpeak uses '?' for glottal stop / stød in some languages (e.g. Danish).
		u"?": (u"ʔ" if u"ʔ" in data else u""),
		# consonants
		u"N": u"ŋ",
		u"T": u"θ",
		u"D": u"ð",
		u"B": u"b",
		u"Q": u"g",
		u"x": (u"x" if isGerman else u"h"),
		# vowels used by some languages
		u"&": u"ɐ",
		u"Y": u"ø",
		u"W": u"œ",
	}

	# Portuguese: eSpeak sometimes uses 'y' for a /j/ glide (e.g. 'mãe' -> &~y).
	# Map it to /j/ so it doesn't turn into a foreign-sounding [y] vowel.
	if isPortuguese:
		asciiMap[u"y"] = u"j"

	# Hungarian long 'á' is /aː/ (distinct from short 'a' /ɒ/).
	# eSpeak Hungarian mnemonics output long á as "a:" (lowercase a + length).
	# We map that lowercase "a" to a dedicated internal vowel symbol so you can
	# tune Hungarian á without affecting Spanish/Portuguese/Polish "a".
	# NOTE: "ᴀ" (small capital A) is used here as an internal placeholder.
	if isHungarian:
		asciiMap[u"a"] = u"ᴀ"

	# A is language-dependent.
	# Hungarian short 'a' is /ɒ/ but we route it to a dedicated internal vowel (ᴒ)
	# so it can be tuned without affecting UK English LOT (/ɒ/).
	if isHungarian:
		asciiMap[u"A"] = u"ᴒ"
	else:
		asciiMap[u"A"] = u"ɑ"

	# S/Z: eSpeak mnemonic phonemes for /ʃ/ and /ʒ/.
	asciiMap[u"S"] = u"ʃ"
	asciiMap[u"Z"] = u"ʒ"

	# Polish frequently uses "R" in mnemonic output for a rolled "r".
	# Map it to our trill/tap phoneme so it doesn't get dropped.
	if isPolish:
		asciiMap[u"R"] = u"r"

	# Portuguese onset clusters: "tr" often outputs R for a tap-like r.
	if isPortuguese:
		asciiMap[u"R"] = u"ɾ"

	# Polish: eSpeak outputs trilled r as "R" and the vowel "y" as /ɨ/.
	if isPolish:
		asciiMap[u"R"] = u"r"
		asciiMap[u"y"] = u"ɨ"

	# Romanian: eSpeak uses 'y' for /ɨ/ (e.g. încă -> 'ynk@).
	if isRomanian:
		asciiMap[u"y"] = u"ɨ"

	# Danish: eSpeak uses 'R' for a uvular/approximant r and '?' for stød/glottal stop.
	if isDanish:
		asciiMap[u"R"] = (u"ʁ" if u"ʁ" in data else u"r")


	# German: eSpeak uses '3' for vocalic -er ("über" -> yb3). Map to ɐ if supported.
	if isGerman and (u"ɐ" in data):
		asciiMap[u"3"] = u"ɐ"
	# German: as a fallback, also map mnemonic 'C' if it survived multi-replacements.
	if isGerman:
		asciiMap[u"C"] = (u"ç" if u"ç" in data else u"x")

	# English LOT vowel differs across accents.
	if isEnglish:
		asciiMap[u"0"] = (u"ɑ" if isRhoticEnglish else u"ɒ")
	else:
		asciiMap[u"0"] = u"ɒ"

	for k, v in asciiMap.items():
		text = text.replace(k, v)

	# English (UK / non-rhotic): if the upstream phoneme stream is already IPA
	# (using 'ɔ' rather than the eSpeak mnemonic 'O'), route it to our rounded
	# THOUGHT vowel ('O') when available. This keeps UK voices from sounding
	# accidentally US/open‑O after you tune 'ɔ' for American English.
	if isEnglish and isNonRhoticEnglish and (u"O" in data):
		text = text.replace(u"ɔ", u"O")

	# Portuguese: keep ó (open O) distinct when we have a dedicated phoneme entry.
	if isPortuguese and (u"ᴐ" in data):
		text = text.replace(u"ɔ", u"ᴐ")

	# Remove leftover mnemonic modifiers.
	text = text.replace(u";", u"")
	text = text.replace(u"^", u"")

	# --- IPA normalisation / fallbacks ---
	# Dark-L and syllabic-L variants.
	text = text.replace(u"l̩", u"əl")
	text = text.replace(u"ɫ̩", u"əl")
	text = text.replace(u"ə͡l", u"əl")
	text = text.replace(u"ə͡l", u"əl")
	text = text.replace(u"ʊ͡l", u"əl")

	# Common reduced/central vowel symbol used by some eSpeak accents.
	if u"ᵻ" not in data:
		text = text.replace(u"ᵻ", u"ɪ")

	# Rhotic hook (˞) and syllabic-r.
	text = text.replace(u"˞", u"ɹ")
	text = text.replace(u"ɹ̩", u"ɚ" if u"ɚ" in data else u"əɹ")
	text = text.replace(u"r̩", u"ɚ" if u"ɚ" in data else u"əɹ")

	# If rhotic vowels don't exist, fall back to vowel+ɹ.
	if u"ɚ" not in data:
		text = text.replace(u"ɚ", u"əɹ")
	if u"ɝ" not in data:
		text = text.replace(u"ɝ", u"ɜɹ")

	# English: normalize 'r' to approximant.
	if isEnglish:
		text = text.replace(u"r", u"ɹ")
	# French/German: use uvular r if supported (these languages usually realise /r/ as [ʁ]).
	if (not isEnglish) and (isFrench or isGerman) and (u"ʁ" in data):
		text = text.replace(u"r", u"ʁ")

	# Cross-language approximations, but only if the target phoneme isn't supported.
	repl = {
		# Polish.
		u"ɕ": (u"ɕ" if u"ɕ" in data else u"ʃ"),
		u"ʑ": (u"ʑ" if u"ʑ" in data else u"ʒ"),
		u"ʂ": (u"ʂ" if u"ʂ" in data else u"ʃ"),
		u"ʐ": (u"ʐ" if u"ʐ" in data else u"ʒ"),
		u"t͡ɕ": (u"t͡ɕ" if u"t͡ɕ" in data else u"t͡ʃ"),
		u"d͡ʑ": (u"d͡ʑ" if u"d͡ʑ" in data else u"d͡ʒ"),
		# Spanish/Portuguese.
		u"β": u"b",
		u"ɣ": u"g",
		u"x": (u"x" if (isGerman and (u"x" in data)) else u"h"),
		u"ʝ": u"j",
		u"ʎ": (u"ʎ" if u"ʎ" in data else u"l"),
		# Palatal stops.
		u"c": u"k",
		u"ɟ": u"g",
		# Nasals.
		u"ɲ": (u"ɲ" if u"ɲ" in data else u"n"),
		# Misc vowels.
		u"ɘ": (u"ɘ" if u"ɘ" in data else u"ə"),
		u"ɵ": (u"ɵ" if u"ɵ" in data else (u"ø" if u"ø" in data else u"o")),
		u"ɤ": (u"ɤ" if u"ɤ" in data else u"ʌ"),
	}
	for k, v in repl.items():
		text = text.replace(k, v)

	# Precomposed nasal vowels.
	# Keep them for Portuguese if we have corresponding phonemes; otherwise fall back.
	if isPortuguese:
		if (u"ã" not in data) or (u"õ" not in data) or (u"ũ" not in data):
			text = text.replace(u"ã", u"a").replace(u"ẽ", u"e").replace(u"ĩ", u"i").replace(u"õ", u"o").replace(u"ũ", u"u")
	else:
		text = text.replace(u"ã", u"a").replace(u"ẽ", u"e").replace(u"ĩ", u"i").replace(u"õ", u"o").replace(u"ũ", u"u")

	# English TRAP: /a/ -> /æ/ in rhotic North American accents.
	if isEnglish:
		text = text.replace(u"a", u"æ")

	# Drop leftover markers.
	text = text.replace(u"#", u"")

	# Collapse whitespace.
	text = re.sub(r"\s+", " ", text).strip()
	return text

def iterPhonemes(**kwargs):
	for k,v in data.items():
		if all(v[x]==y for x,y in kwargs.items()):
			yield k

def setFrame(frame,phoneme):
	values=data[phoneme]
	for k,v in values.items():
		setattr(frame,k,v)

def applyPhonemeToFrame(frame,phoneme):
	for k,v in phoneme.items():
		if not k.startswith('_'):
			setattr(frame,k,v)

def _IPAToPhonemesHelper(text):
	textLen=len(text)
	index=0
	offset=0
	curStress=0
	for index in range(textLen):
		index=index+offset
		if index>=textLen:
			break
		char=text[index]
		if char=='ˈ':
			curStress=1
			continue
		elif char=='ˌ':
			curStress=2
			continue
		isLengthened=(text[index+1:index+2]=='ː')
		isTiedTo=(text[index+1:index+2]=='͡')
		isTiedFrom=(text[index-1:index]=='͡') if index>0 else False
		phoneme=None
		if isTiedTo:
			phoneme=data.get(text[index:index+3])
			offset+=2 if phoneme else 1
		elif isLengthened:
			phoneme=data.get(text[index:index+2])
			offset+=1
		if not phoneme:
			phoneme=data.get(char)
		if not phoneme:
			yield char,None
			continue
		phoneme=phoneme.copy()
		if curStress:
			phoneme['_stress']=curStress
			curStress=0
		if isTiedFrom:
			phoneme['_tiedFrom']=True
		elif isTiedTo:
			phoneme['_tiedTo']=True
		if isLengthened:
			phoneme['_lengthened']=True
		phoneme['_char']=char
		yield char,phoneme

def IPAToPhonemes(ipaText, language=None):
	phonemeList=[]
	textLength=len(ipaText)
	lang=(language or '').replace('_','-').lower()
	# Default to English behaviour if the caller didn't provide a language.
	isEnglish=(not lang) or lang.startswith('en')
	# Collect phoneme info for each IPA character, assigning diacritics (lengthened, stress) to the last real phoneme
	newWord=True
	lastPhoneme=None
	syllableStartPhoneme=None
	for char,phoneme in _IPAToPhonemesHelper(ipaText):
		if char==' ':
			newWord=True
		elif phoneme:
			stress=phoneme.pop('_stress',0)
			if lastPhoneme and not lastPhoneme.get('_isVowel') and phoneme and phoneme.get('_isVowel'):
				lastPhoneme['_syllableStart']=True
				syllableStartPhoneme=lastPhoneme
			elif stress==1 and lastPhoneme and lastPhoneme.get('_isVowel'):
				phoneme['_syllableStart']=True
				syllableStartPhoneme=phoneme
			if isEnglish and lastPhoneme and lastPhoneme.get('_isStop') and not lastPhoneme.get('_isVoiced') and phoneme and phoneme.get('_isVoiced') and not phoneme.get('_isStop') and not phoneme.get('_isAfricate'): 
				psa=data['h'].copy()
				psa['_postStopAspiration']=True
				psa['_char']=None
				phonemeList.append(psa)
				lastPhoneme=psa
			if newWord:
				newWord=False
				phoneme['_wordStart']=True
				phoneme['_syllableStart']=True
				syllableStartPhoneme=phoneme
			if stress:
				syllableStartPhoneme['_stress']=stress
			elif phoneme.get('_isStop') or phoneme.get('_isAfricate'):
				gap=dict(_silence=True,_preStopGap=True)
				phonemeList.append(gap)
			phonemeList.append(phoneme)
			lastPhoneme=phoneme
	return phonemeList

def correctHPhonemes(phonemeList):
	finalPhonemeIndex=len(phonemeList)-1
	# Correct all h phonemes (including inserted aspirations) so that their formants match the next phoneme, or the previous if there is no next
	for index in range(len(phonemeList)):
		prevPhoneme=phonemeList[index-1] if index>0 else None
		curPhoneme=phonemeList[index]
		nextPhoneme=phonemeList[index+1] if index<finalPhonemeIndex else None
		if curPhoneme.get('_copyAdjacent'):
			adjacent=nextPhoneme if nextPhoneme and not nextPhoneme.get('_silence') else prevPhoneme 
			if adjacent:
				for k,v in adjacent.items():
					if not k.startswith('_') and k not in curPhoneme:
						curPhoneme[k]=v

def calculatePhonemeTimes(phonemeList,baseSpeed,language=None):
	lastPhoneme=None
	syllableStress=0
	speed=baseSpeed
	lang=(language or '').replace('_','-').lower()
	primaryStressDiv, secondaryStressDiv = _get_stress_slowdown(language)
	isEnglish=lang.startswith('en')
	# Keep rhotic/non-rhotic detection consistent with normalizeIPA:
	# treat English as rhotic by default unless explicitly en-gb/en-uk.
	# Keep rhotic/non-rhotic classification consistent with normalizeIPA.
	# eSpeak's "en" is British, so treat it as non-rhotic.
	isNonRhoticEnglish=(isEnglish and (lang == 'en' or lang.startswith('en-gb') or lang.startswith('en-uk')))
	isRhoticEnglish=(isEnglish and not isNonRhoticEnglish)
	isPortuguese=lang.startswith('pt')
	for index,phoneme in enumerate(phonemeList):
		nextPhoneme=phonemeList[index+1] if len(phonemeList)>index+1 else None
		syllableStart=phoneme.get('_syllableStart')
		if syllableStart:
			syllableStress=phoneme.get('_stress')
			if syllableStress:
				# Slow down stressed syllables to shape cadence.
				speed=baseSpeed/primaryStressDiv if syllableStress==1 else baseSpeed/secondaryStressDiv
			else:
				speed=baseSpeed
		phonemeDuration=60.0/speed
		phonemeFadeDuration=10.0/speed
		if phoneme.get('_preStopGap'):
			phonemeDuration=41.0/speed
		elif phoneme.get('_postStopAspiration'):
			phonemeDuration=20.0/speed
		elif phoneme.get('_isTap') or phoneme.get('_isTrill'):
			# Alveolar tap/trill: keep it short, but don't force a silence gap like a full stop.
			if phoneme.get('_isTrill'):
				phonemeDuration=22.0/speed
			else:
				phonemeDuration=min(14.0/speed,14.0)
			phonemeFadeDuration=0.001
		elif phoneme.get('_isStop'):
			phonemeDuration=min(6.0/speed,6.0)
			phonemeFadeDuration=0.001
		elif phoneme.get('_isAfricate'):
			phonemeDuration=24.0/speed
			phonemeFadeDuration=0.001
		elif not phoneme.get('_isVoiced'):
			phonemeDuration=45.0/speed
		else: # is voiced
			if phoneme.get('_isVowel'):
				if lastPhoneme and (lastPhoneme.get('_isLiquid') or lastPhoneme.get('_isSemivowel')): 
					phonemeFadeDuration=25.0/speed
				if phoneme.get('_tiedTo'):
					# English PRICE/MOUTH (aɪ/aʊ) can sound too clipped;
					# give the first element a touch more time without affecting FACE/GOAT.
					if isEnglish and phoneme.get('_char') == u'ɑ':
						phonemeDuration=42.0/speed
					else:
						phonemeDuration=40.0/speed
				elif phoneme.get('_tiedFrom'):
					# Make the offglide a bit more audible for aɪ/aʊ (five/nine).
					if (isEnglish and phoneme.get('_char') in (u'ɪ', u'ʊ')
						and lastPhoneme and lastPhoneme.get('_tiedTo') and lastPhoneme.get('_char') == u'ɑ'):
						phonemeDuration=24.0/speed
						phonemeFadeDuration=18.0/speed
					else:
						phonemeDuration=20.0/speed
						phonemeFadeDuration=20.0/speed
				elif not syllableStress and not syllableStart and nextPhoneme and not nextPhoneme.get('_wordStart') and (nextPhoneme.get('_isLiquid') or nextPhoneme.get('_isNasal')):
					if nextPhoneme.get('_isLiquid'):
						phonemeDuration=30.0/speed
					else:
						phonemeDuration=40.0/speed
			else: # not a vowel
				phonemeDuration=30.0/speed
				if phoneme.get('_isLiquid') or phoneme.get('_isSemivowel'):
					phonemeFadeDuration=20.0/speed

		# Hungarian short "a" (A -> ᴒ) should be noticeably shorter than long "á".
		# We already lengthen long vowels via the _lengthened path; here we slightly
		# compress the short vowel so it doesn't drift toward "á" in running speech.
		if lang.startswith('hu') and phoneme.get('_isVowel') and phoneme.get('_char') == u'ᴒ' and not phoneme.get('_lengthened'):
			phonemeDuration *= 0.85

		# English word-final long /uː/ (blue, new, view) can sound over-held,
		# especially after liquids/semivowels (blue, view).
		# Shorten it slightly and reduce the fade so it doesn't 'hang' at the end.
		if isEnglish and phoneme.get('_isVowel') and phoneme.get('_char') == u'u' and phoneme.get('_lengthened'):
			if (nextPhoneme is None) or nextPhoneme.get('_wordStart'):
				phonemeDuration *= 0.80
				phonemeFadeDuration = min(phonemeFadeDuration, 14.0/speed)
		if phoneme.get('_lengthened'):
			# Vowel length is phonemic in Hungarian; make it clearly longer.
			lang=(language or '').replace('_','-').lower()
			if lang.startswith('hu'):
				phonemeDuration*=1.3
			else:
				phonemeDuration*=1.05

		phoneme['_duration']=phonemeDuration
		phoneme['_fadeDuration']=phonemeFadeDuration
		lastPhoneme=phoneme

def applyPitchPath(phonemeList,startIndex,endIndex,basePitch,inflection,startPitchPercent,endPitchPercent):
	startPitch=basePitch*(2**(((startPitchPercent-50)/50.0)*inflection))
	endPitch=basePitch*(2**(((endPitchPercent-50)/50.0)*inflection))
	voicedDuration=0
	for index in range(startIndex,endIndex):
		phoneme=phonemeList[index]
		if phoneme.get('_isVoiced'):
			voicedDuration+=phoneme['_duration']
	curDuration=0
	pitchDelta=endPitch-startPitch
	curPitch=startPitch
	syllableStress=False
	for index in range(startIndex,endIndex):
		phoneme=phonemeList[index]
		phoneme['voicePitch']=curPitch
		if phoneme.get('_isVoiced'):
			curDuration+=phoneme['_duration']
			pitchRatio=curDuration/float(voicedDuration)
			curPitch=startPitch+(pitchDelta*pitchRatio)
		phoneme['endVoicePitch']=curPitch

intonationParamTable={
	'.':{
		'preHeadStart':46,
		'preHeadEnd':57,
		'headExtendFrom':4,
		'headStart':80,
		'headEnd':50,
		'headSteps':[100,75,50,25,0,63,38,13,0],
		'headStressEndDelta':-16,
		'headUnstressedRunStartDelta':-8,
		'headUnstressedRunEndDelta':-5,
		'nucleus0Start':64,
		'nucleus0End':8,
		'nucleusStart':70,
		'nucleusEnd':18,
		'tailStart':24,
		'tailEnd':8,
	},
	',':{
		'preHeadStart':46,
		'preHeadEnd':57,
		'headExtendFrom':4,
		'headStart':80,
		'headEnd':60,
		'headSteps':[100,75,50,25,0,63,38,13,0],
		'headStressEndDelta':-16,
		'headUnstressedRunStartDelta':-8,
		'headUnstressedRunEndDelta':-5,
		'nucleus0Start':34,
		'nucleus0End':52,
		'nucleusStart':78,
		'nucleusEnd':34,
		'tailStart':34,
		'tailEnd':52,
	},
	'?':{
		'preHeadStart':45,
		'preHeadEnd':56,
		'headExtendFrom':3,
		'headStart':75,
		'headEnd':43,
		'headSteps':[100,75,50,20,60,35,11,0],
		'headStressEndDelta':-16,
		'headUnstressedRunStartDelta':-7,
		'headUnstressedRunEndDelta':0,
		'nucleus0Start':34,
		'nucleus0End':68,
		'nucleusStart':86,
		'nucleusEnd':21,
		'tailStart':34,
		'tailEnd':68,
	},
	'!':{
		'preHeadStart':46,
		'preHeadEnd':57,
		'headExtendFrom':3,
		'headStart':90,
		'headEnd':50,
		'headSteps':[100,75,50,16,82,50,32,16],
		'headStressEndDelta':-16,
		'headUnstressedRunStartDelta':-9,
		'headUnstressedRunEndDelta':0,
		'nucleus0Start':92,
		'nucleus0End':4,
		'nucleusStart':92,
		'nucleusEnd':80,
		'tailStart':76,
		'tailEnd':4,
	}
}

def calculatePhonemePitches(phonemeList,speed,basePitch,inflection,clauseType):
	intonationParams=intonationParamTable[clauseType or '.']
	preHeadStart=0
	preHeadEnd=len(phonemeList)
	for index,phoneme in enumerate(phonemeList):
		if phoneme.get('_syllableStart'):
			syllableStress=phoneme.get('_stress')==1
			if syllableStress:
				preHeadEnd=index
				break
	if (preHeadEnd-preHeadStart)>0:
		applyPitchPath(phonemeList,preHeadStart,preHeadEnd,basePitch,inflection,intonationParams['preHeadStart'],intonationParams['preHeadEnd'])
	nucleusStart=nucleusEnd=tailStart=tailEnd=len(phonemeList)
	for index in range(nucleusEnd-1,preHeadEnd-1,-1):
		phoneme=phonemeList[index]
		if phoneme.get('_syllableStart'):
			syllableStress=phoneme.get('_stress')==1
			if syllableStress :
				nucleusStart=index
				break
			else:
				nucleusEnd=tailStart=index
	hasTail=(tailEnd-tailStart)>0
	if hasTail:
		applyPitchPath(phonemeList,tailStart,tailEnd,basePitch,inflection,intonationParams['tailStart'],intonationParams['tailEnd'])
	if (nucleusEnd-nucleusStart)>0:
		if hasTail:
			applyPitchPath(phonemeList,nucleusStart,nucleusEnd,basePitch,inflection,intonationParams['nucleusStart'],intonationParams['nucleusEnd'])
		else:
			applyPitchPath(phonemeList,nucleusStart,nucleusEnd,basePitch,inflection,intonationParams['nucleus0Start'],intonationParams['nucleus0End'])
	if preHeadEnd<nucleusStart:
		headStartPitch=intonationParams['headStart']
		headEndPitch=intonationParams['headEnd']
		lastHeadStressStart=None
		lastHeadUnstressedRunStart=None
		stressEndPitch=None
		steps=intonationParams['headSteps']
		extendFrom=intonationParams['headExtendFrom']
		stressStartPercentageGen=itertools.chain(steps,itertools.cycle(steps[extendFrom:]))
		for index in range(preHeadEnd,nucleusStart+1):
			phoneme=phonemeList[index]
			syllableStress=phoneme.get('_stress')==1
			if phoneme.get('_syllableStart'):
				if lastHeadStressStart is not None:
					stressStartPitch=headEndPitch+(((headStartPitch-headEndPitch)/100.0)*next(stressStartPercentageGen))
					stressEndPitch=stressStartPitch+intonationParams['headStressEndDelta']
					applyPitchPath(phonemeList,lastHeadStressStart,index,basePitch,inflection,stressStartPitch,stressEndPitch)
					lastHeadStressStart=None
				if syllableStress :
					if lastHeadUnstressedRunStart is not None:
						unstressedRunStartPitch=stressEndPitch+intonationParams['headUnstressedRunStartDelta']
						unstressedRunEndPitch=stressEndPitch+intonationParams['headUnstressedRunEndDelta']
						applyPitchPath(phonemeList,lastHeadUnstressedRunStart,index,basePitch,inflection,unstressedRunStartPitch,unstressedRunEndPitch)
						lastHeadUnstressedRunStart=None
					lastHeadStressStart=index
				elif lastHeadUnstressedRunStart is None: 
					lastHeadUnstressedRunStart=index

def generateFramesAndTiming(ipaText,speed=1,basePitch=100,inflection=0.5,clauseType=None,language=None):
	ipaText=normalizeIPA(ipaText,language=language)
	phonemeList=IPAToPhonemes(ipaText, language=language)
	if len(phonemeList)==0:
		return
	correctHPhonemes(phonemeList)
	calculatePhonemeTimes(phonemeList,speed,language=language)
	calculatePhonemePitches(phonemeList,speed,basePitch,inflection,clauseType)
	for phoneme in phonemeList:
		frameDuration=phoneme.pop('_duration')
		fadeDuration=phoneme.pop('_fadeDuration')
		if phoneme.get('_silence'):
			yield None,frameDuration,fadeDuration
		else:
			frame=speechPlayer.Frame()
			frame.preFormantGain=1.0
			frame.outputGain=1.5
			applyPhonemeToFrame(frame,phoneme)
			yield frame,frameDuration,fadeDuration
