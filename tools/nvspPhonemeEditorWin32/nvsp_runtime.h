#pragma once

#include <string>
#include <vector>

#include <windows.h>

#include "speechPlayer.h"
#include "yaml_edit.h"

// NOTE: We dynamically load speechPlayer.dll and nvspFrontend.dll.
// These types are used only for layout compatibility.
#include "nvspFrontend.h"

namespace nvsp_editor {

// -------------------------
// Dynamic DLL function types
// -------------------------

// speechPlayer.dll API
using sp_initialize_fn = speechPlayer_handle_t(*)(int);
using sp_queueFrame_fn = void(*)(speechPlayer_handle_t, speechPlayer_frame_t*, unsigned int, unsigned int, int, bool);
using sp_synthesize_fn = int(*)(speechPlayer_handle_t, unsigned int, sample*);
using sp_terminate_fn = void(*)(speechPlayer_handle_t);

// nvspFrontend.dll API
using fe_create_fn = nvspFrontend_handle_t(*)(const char*);
using fe_destroy_fn = void(*)(nvspFrontend_handle_t);
using fe_setLanguage_fn = int(*)(nvspFrontend_handle_t, const char*);
using fe_queueIPA_fn = int(*)(
  nvspFrontend_handle_t,
  const char*,
  double,
  double,
  double,
  const char*,
  int,
  nvspFrontend_FrameCallback,
  void*
);
using fe_getLastError_fn = const char*(*)(nvspFrontend_handle_t);

class NvspRuntime {
public:
  NvspRuntime();
  ~NvspRuntime();

  // Directory containing speechPlayer.dll and nvspFrontend.dll.
  bool setDllDirectory(const std::wstring& dllDir, std::string& outError);

  // Directory that contains a "packs" folder.
  bool setPackRoot(const std::wstring& packRootDir, std::string& outError);

  // Language tag like "en-us", "hu", ...
  bool setLanguage(const std::string& langTagUtf8, std::string& outError);

  bool dllsLoaded() const;

  // Synthesize just a single phoneme (from phonemes.yaml) to PCM samples.
  bool synthPreviewPhoneme(
    const Node& phonemeMap,
    int sampleRate,
    std::vector<sample>& outSamples,
    std::string& outError
  );

  // Synthesize an IPA string via nvspFrontend.dll to PCM samples.
  bool synthIpa(
    const std::string& ipaUtf8,
    int sampleRate,
    std::vector<sample>& outSamples,
    std::string& outError
  );

  // Last frontend error (if available).
  std::string lastFrontendError() const { return m_lastFrontendError; }

private:
  void unload();

  // DLL modules.
  HMODULE m_speechPlayer = NULL;
  HMODULE m_frontend = NULL;

  sp_initialize_fn m_spInitialize = nullptr;
  sp_queueFrame_fn m_spQueueFrame = nullptr;
  sp_synthesize_fn m_spSynthesize = nullptr;
  sp_terminate_fn m_spTerminate = nullptr;

  fe_create_fn m_feCreate = nullptr;
  fe_destroy_fn m_feDestroy = nullptr;
  fe_setLanguage_fn m_feSetLanguage = nullptr;
  fe_queueIPA_fn m_feQueueIPA = nullptr;
  fe_getLastError_fn m_feGetLastError = nullptr;

  // Runtime state
  nvspFrontend_handle_t m_feHandle = nullptr;
  std::string m_lastFrontendError;
  std::wstring m_packRoot;
  std::string m_langTag;
};

} // namespace nvsp_editor
