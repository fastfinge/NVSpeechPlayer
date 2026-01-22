#pragma once

#define UNICODE
#define _UNICODE

#include <windows.h>

#include <string>
#include <utility>
#include <vector>

#include "nvsp_runtime.h"
#include "yaml_edit.h"

struct AddMappingDialogState {
  nvsp_editor::ReplacementRule rule;
  std::vector<std::string> classNames;
  bool ok = false;
};

struct ClonePhonemeDialogState {
  std::vector<std::string> keys;
  std::string fromKey;
  std::string newKey;
  bool ok = false;
};

struct EditValueDialogState {
  std::string field;
  std::string value;
  nvsp_editor::Node baseMap;
  nvsp_editor::NvspRuntime* runtime = nullptr;

  bool livePreview = true;
  bool ok = false;

  bool armed = false;
  UINT_PTR previewTimer = 0;
};

struct EditSettingDialogState {
  std::string key;
  std::string value;
  std::vector<std::string> knownKeys;
  bool ok = false;
};

struct EditSettingsDialogState {
  std::vector<std::pair<std::string, std::string>> settings;  // key/value
  std::vector<std::string> knownKeys;
  bool ok = false;
};

struct EditPhonemeDialogState {
  std::string phonemeKey;
  nvsp_editor::Node original;
  nvsp_editor::Node working;
  nvsp_editor::NvspRuntime* runtime = nullptr;
  bool ok = false;
};

struct SpeechSettingsDialogState {
  nvsp_editor::SpeechSettings settings;
  nvsp_editor::NvspRuntime* runtime = nullptr;

  // Param UI
  std::vector<std::string> paramNames;
  int selectedParam = 0;
  bool ok = false;
};

// Dialog launch helpers (return true if OK was pressed).
bool ShowAddMappingDialog(HINSTANCE hInst, HWND parent, AddMappingDialogState& st);
bool ShowClonePhonemeDialog(HINSTANCE hInst, HWND parent, ClonePhonemeDialogState& st);
bool ShowEditValueDialog(HINSTANCE hInst, HWND parent, EditValueDialogState& st);
bool ShowEditSettingsDialog(HINSTANCE hInst, HWND parent, EditSettingsDialogState& st);
bool ShowEditPhonemeDialog(HINSTANCE hInst, HWND parent, EditPhonemeDialogState& st);
bool ShowSpeechSettingsDialog(HINSTANCE hInst, HWND parent, SpeechSettingsDialogState& st);

// Persistence for speech settings (nvspPhonemeEditor.ini).
nvsp_editor::SpeechSettings loadSpeechSettingsFromIni();
void saveSpeechSettingsToIni(const nvsp_editor::SpeechSettings& s);
