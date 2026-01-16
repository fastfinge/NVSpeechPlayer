/*
This file is a part of the NV Speech Player project. 
URL: https://bitbucket.org/nvaccess/speechplayer
Copyright 2014 NV Access Limited.
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 2.0, as published by
the Free Software Foundation.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
This license can be found at:
http://www.gnu.org/licenses/old-licenses/gpl-2.0.html
*/

#include <queue>
#include <cmath>
#include <cstring>
#include "utils.h"
#include "frame.h"

using namespace std;

static inline bool isFinite(double v) {
	return std::isfinite(v) != 0;
}

static inline bool looksLikePitchHz(double hz) {
	// Eloquence-style F0 is typically tens to a few hundred Hz,
	// but we keep the window wide to be tolerant.
	return isFinite(hz) && hz >= 20.0 && hz <= 2000.0;
}

// Compatibility helper:
// Newer builds added outputGain immediately before endVoicePitch.
// If an older caller still writes the old struct layout, what we see as
// outputGain may actually be the caller's endVoicePitch, while endVoicePitch
// itself becomes garbage.
//
// This function fixes that *only* when endVoicePitch is clearly invalid and
// outputGain looks like a pitch in Hz.
static inline void normalizeFrameForCompat(speechPlayer_frame_t& f) {
	// Ensure gain is finite.
	if(!isFinite(f.outputGain)) {
		f.outputGain = 1.0;
	}

	const bool endPitchOk = looksLikePitchHz(f.endVoicePitch);
	const bool outGainLooksLikePitch = looksLikePitchHz(f.outputGain) && f.outputGain > 10.0;

	if(!endPitchOk && outGainLooksLikePitch) {
		f.endVoicePitch = f.outputGain;
		f.outputGain = 1.0;
	}

	// If endVoicePitch is missing/invalid, don't ramp pitch to 0.
	if(!looksLikePitchHz(f.endVoicePitch) && looksLikePitchHz(f.voicePitch)) {
		f.endVoicePitch = f.voicePitch;
	}

	// Keep outputGain finite and non-negative.
	// Some callers use >1.0 as a loudness boost, so we allow a reasonable range.
	if(!isFinite(f.outputGain) || f.outputGain < 0.0) f.outputGain = 0.0;
	if(f.outputGain > 8.0) f.outputGain = 8.0;

	// Guard voicePitch.
	if(!isFinite(f.voicePitch) || f.voicePitch < 0.0) {
		f.voicePitch = 0.0;
		f.endVoicePitch = 0.0;
	}
}

struct frameRequest_t {
	unsigned int minNumSamples;
	unsigned int numFadeSamples;
	bool NULLFrame;
	speechPlayer_frame_t frame;
	double voicePitchInc; 
	int userIndex;
};

class FrameManagerImpl: public FrameManager {
	private:
	LockableObject frameLock;
	queue<frameRequest_t*> frameRequestQueue;
	frameRequest_t* oldFrameRequest;
	frameRequest_t* newFrameRequest;
	speechPlayer_frame_t curFrame;
	bool curFrameIsNULL;
	unsigned int sampleCounter;
	int lastUserIndex;
	unsigned int silenceTailSamplesRemaining;
	static const unsigned int SILENCE_TAIL_SAMPLES = 256;

	void updateCurrentFrame() {
		sampleCounter++;
		if(newFrameRequest) {
			if(sampleCounter>(newFrameRequest->numFadeSamples)) {
				delete oldFrameRequest;
				oldFrameRequest=newFrameRequest;
				newFrameRequest=NULL;
			} else {
				double curFadeRatio=(double)sampleCounter/(newFrameRequest->numFadeSamples);
				for(int i=0;i<speechPlayer_frame_numParams;++i) {
					((speechPlayer_frameParam_t*)&curFrame)[i]=calculateValueAtFadePosition(((speechPlayer_frameParam_t*)&(oldFrameRequest->frame))[i],((speechPlayer_frameParam_t*)&(newFrameRequest->frame))[i],curFadeRatio);
				}
			}
		} else if(sampleCounter>(oldFrameRequest->minNumSamples)) {
			if(!frameRequestQueue.empty()) {
				curFrameIsNULL=false;
				silenceTailSamplesRemaining=0;
				newFrameRequest=frameRequestQueue.front();
				frameRequestQueue.pop();
				if(newFrameRequest->NULLFrame) {
					memcpy(&(newFrameRequest->frame),&(oldFrameRequest->frame),sizeof(speechPlayer_frame_t));
					newFrameRequest->frame.preFormantGain=0;
					newFrameRequest->frame.voicePitch=curFrame.voicePitch;
					newFrameRequest->voicePitchInc=0;
				} else if(oldFrameRequest->NULLFrame) {
					memcpy(&(oldFrameRequest->frame),&(newFrameRequest->frame),sizeof(speechPlayer_frame_t));
					oldFrameRequest->frame.preFormantGain=0;
				}
				if(newFrameRequest) {
					if(newFrameRequest->userIndex!=-1) lastUserIndex=newFrameRequest->userIndex;
					sampleCounter=0;
					newFrameRequest->frame.voicePitch+=(newFrameRequest->voicePitchInc*newFrameRequest->numFadeSamples);
				}
			} else {
				// Queue drained. Keep returning a silent frame briefly so the
				// resonators can ring down, then return NULL to signal end-of-stream.
				if(!curFrameIsNULL) {
					curFrame.preFormantGain=0;
					oldFrameRequest->frame.preFormantGain=0;
					if(silenceTailSamplesRemaining==0) {
						// +1 so we output exactly SILENCE_TAIL_SAMPLES samples before switching to NULL.
						silenceTailSamplesRemaining=SILENCE_TAIL_SAMPLES+1;
					}
					if(silenceTailSamplesRemaining>0) {
						--silenceTailSamplesRemaining;
						if(silenceTailSamplesRemaining==0) curFrameIsNULL=true;
					}
				}
			}
		} else {
			curFrame.voicePitch+=oldFrameRequest->voicePitchInc;
			oldFrameRequest->frame.voicePitch=curFrame.voicePitch;
		}
	}


	public:

	FrameManagerImpl(): curFrame(), curFrameIsNULL(true), sampleCounter(0), silenceTailSamplesRemaining(0), newFrameRequest(NULL), lastUserIndex(-1)  {
		oldFrameRequest=new frameRequest_t();
		oldFrameRequest->NULLFrame=true;
	}

	void queueFrame(speechPlayer_frame_t* frame, unsigned int minNumSamples, unsigned int numFadeSamples, int userIndex, bool purgeQueue) {
		frameLock.acquire();
		frameRequest_t* frameRequest=new frameRequest_t();
		frameRequest->minNumSamples=(minNumSamples<1u)?1u:minNumSamples;
		frameRequest->numFadeSamples=(numFadeSamples<1u)?1u:numFadeSamples;
		if(frame) {
			frameRequest->NULLFrame=false;
			memcpy(&(frameRequest->frame),frame,sizeof(speechPlayer_frame_t));
			normalizeFrameForCompat(frameRequest->frame);
			frameRequest->voicePitchInc=(frameRequest->frame.endVoicePitch-frameRequest->frame.voicePitch)/frameRequest->minNumSamples;
		} else {
			frameRequest->NULLFrame=true;
		}
		frameRequest->userIndex=userIndex;
		if(purgeQueue) {
			for(;!frameRequestQueue.empty();frameRequestQueue.pop()) delete frameRequestQueue.front();
			sampleCounter=oldFrameRequest->minNumSamples;
			silenceTailSamplesRemaining=0;
			if(newFrameRequest) {
				oldFrameRequest->NULLFrame=newFrameRequest->NULLFrame;
				memcpy(&(oldFrameRequest->frame),&curFrame,sizeof(speechPlayer_frame_t));
				delete newFrameRequest;
				newFrameRequest=NULL;
			}
		}
		frameRequestQueue.push(frameRequest);
		frameLock.release();
	}

	const int getLastIndex() {
		return lastUserIndex;
	}

	const speechPlayer_frame_t* const getCurrentFrame() {
		frameLock.acquire();
		updateCurrentFrame();
		frameLock.release();
		return curFrameIsNULL?NULL:&curFrame;
	}

	~FrameManagerImpl() {
		for(;!frameRequestQueue.empty();frameRequestQueue.pop()) delete frameRequestQueue.front();
		if(oldFrameRequest) delete oldFrameRequest;
		if(newFrameRequest) delete newFrameRequest;
	}

};

FrameManager* FrameManager::create() { return new FrameManagerImpl(); }
