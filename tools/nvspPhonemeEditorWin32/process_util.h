#pragma once

#include <string>
#include <vector>

#include <windows.h>

namespace nvsp_editor {

// Run a process and capture its stdout as UTF-8 bytes.
// exePath: full path to the exe.
// args: command line arguments (without the exe name).
// Returns true on success.
bool runProcessCaptureStdout(
  const std::wstring& exePath,
  const std::wstring& args,
  std::string& outStdoutUtf8,
  std::string& outError
);

// Find espeak-ng.exe or espeak.exe inside a directory.
std::wstring findEspeakExe(const std::wstring& espeakDir);

} // namespace nvsp_editor
