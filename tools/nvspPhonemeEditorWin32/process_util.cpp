#include "process_util.h"

#include <filesystem>
#include <sstream>

namespace fs = std::filesystem;

namespace nvsp_editor {

static std::wstring quoteArg(const std::wstring& s) {
  // Simple quoting for CreateProcess command lines.
  if (s.empty()) return L"\"\"";
  bool needs = false;
  for (wchar_t c : s) {
    if (c == L' ' || c == L'\t' || c == L'\n' || c == L'\v' || c == L'"') {
      needs = true;
      break;
    }
  }
  if (!needs) return s;

  std::wstring out = L"\"";
  for (wchar_t c : s) {
    if (c == L'"') out += L"\\\"";
    else out.push_back(c);
  }
  out += L"\"";
  return out;
}

bool runProcessCaptureStdout(
  const std::wstring& exePath,
  const std::wstring& args,
  std::string& outStdoutUtf8,
  std::string& outError
) {
  outStdoutUtf8.clear();
  outError.clear();

  if (exePath.empty()) {
    outError = "Executable path is empty";
    return false;
  }

  HANDLE hRead = NULL;
  HANDLE hWrite = NULL;

  SECURITY_ATTRIBUTES sa{};
  sa.nLength = sizeof(sa);
  sa.bInheritHandle = TRUE;

  if (!CreatePipe(&hRead, &hWrite, &sa, 0)) {
    outError = "CreatePipe failed";
    return false;
  }

  // Ensure the read handle is not inherited.
  SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0);

  STARTUPINFOW si{};
  si.cb = sizeof(si);
  si.dwFlags = STARTF_USESTDHANDLES;
  si.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
  si.hStdOutput = hWrite;
  si.hStdError = hWrite;

  PROCESS_INFORMATION pi{};

  std::wstring cmd = quoteArg(exePath);
  if (!args.empty()) {
    cmd += L" ";
    cmd += args;
  }

  // CreateProcess wants a writable buffer.
  std::vector<wchar_t> cmdBuf(cmd.begin(), cmd.end());
  cmdBuf.push_back(L'\0');

  BOOL ok = CreateProcessW(
    exePath.c_str(),
    cmdBuf.data(),
    NULL,
    NULL,
    TRUE,
    CREATE_NO_WINDOW,
    NULL,
    NULL,
    &si,
    &pi
  );

  // Parent doesn't write.
  CloseHandle(hWrite);
  hWrite = NULL;

  if (!ok) {
    DWORD e = GetLastError();
    CloseHandle(hRead);
    std::ostringstream oss;
    oss << "CreateProcess failed (" << static_cast<unsigned long>(e) << ")";
    outError = oss.str();
    return false;
  }

  // Read all output.
  std::string buf;
  char tmp[4096];
  DWORD read = 0;
  while (ReadFile(hRead, tmp, sizeof(tmp), &read, NULL) && read > 0) {
    buf.append(tmp, tmp + read);
  }

  CloseHandle(hRead);
  hRead = NULL;

  WaitForSingleObject(pi.hProcess, INFINITE);

  DWORD exitCode = 0;
  GetExitCodeProcess(pi.hProcess, &exitCode);

  CloseHandle(pi.hThread);
  CloseHandle(pi.hProcess);

  outStdoutUtf8 = std::move(buf);

  // Trim trailing CR/LF.
  while (!outStdoutUtf8.empty() && (outStdoutUtf8.back() == '\n' || outStdoutUtf8.back() == '\r')) {
    outStdoutUtf8.pop_back();
  }

  if (exitCode != 0 && outStdoutUtf8.empty()) {
    std::ostringstream oss;
    oss << "Process exit code " << static_cast<unsigned long>(exitCode);
    outError = oss.str();
    return false;
  }

  return true;
}

std::wstring findEspeakExe(const std::wstring& espeakDir) {
  if (espeakDir.empty()) return L"";

  fs::path base(espeakDir);

  fs::path ng = base / "espeak-ng.exe";
  if (fs::exists(ng)) return ng.wstring();

  fs::path legacy = base / "espeak.exe";
  if (fs::exists(legacy)) return legacy.wstring();

  return L"";
}

} // namespace nvsp_editor
