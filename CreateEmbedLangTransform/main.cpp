#include <Windows.h>
#include <tchar.h>
#include <MsiQuery.h>
#include <iostream>
#include <string>

#pragma comment(lib, "msi.lib")

#ifdef _UNICODE
#define tcerr wcerr
#else
#define tcerr cerr
#endif // _UNICODE
using tstring = std::basic_string<TCHAR>;

void HandleWindowsError(LPCTSTR func) {
  DWORD error{GetLastError()};
  if (error == ERROR_SUCCESS) return;
  LPTSTR errorMessage;
  FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, error, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), reinterpret_cast<LPTSTR>(&errorMessage), 0u, nullptr);
  std::tcerr << TEXT("Windows error at ") << func << TEXT(": ") << errorMessage << TEXT(" (") << error << TEXT(")") << std::endl;
  LocalFree(errorMessage);
}

void HandleMsiError(UINT result, LPCTSTR func) {
  PMSIHANDLE error{MsiGetLastErrorRecord()};
  if (error) {
    DWORD length{0u};
    auto status{MsiFormatRecord(0u, error, TEXT(""), &length)}; // get the length of the error message
    if (status == ERROR_MORE_DATA) {
      tstring errorMessage;
      errorMessage.resize(length);
      length++; // for the '\0' terminator
      status = MsiFormatRecord(0u, error, errorMessage.data(), &length);
      if (status == ERROR_SUCCESS) {
        std::tcerr << TEXT("MSI error at ") << func << TEXT(": ") << errorMessage << TEXT(" (") << result << TEXT(")") << std::endl;
      }
    }
  }
  else { // no error record, just print the error code
    std::tcerr << TEXT("MSI error at ") << func << TEXT(" (") << result << TEXT(")") << std::endl;
    HandleWindowsError(func);
  }
}

int _tmain(int argc, TCHAR **argv) {
  if (argc < 4) {
    std::tcerr << TEXT("ERROR: too few arguments") << std::endl;
    std::tcerr << TEXT("usage: ") << argv[0] << TEXT(" <target-and-reference-msi> <source-msi> <language-identifier>") << std::endl;
    return 1;
  }

  PMSIHANDLE source, target, view, record, summaryInfo;
  UINT uResult, dataType;
  DWORD dwResult, length{0u};
  BOOL bResult;
  TCHAR tempPath[MAX_PATH], tempFile[MAX_PATH];

  dwResult = GetTempPath(MAX_PATH, tempPath);
  if (dwResult > MAX_PATH || dwResult == 0u) {
    HandleWindowsError(TEXT("GetTempPath"));
    return 1;
  }
  uResult = GetTempFileName(tempPath, TEXT("CELT"), 0u, tempFile);
  if (uResult == 0) {
    HandleWindowsError(TEXT("GetTempFileName"));
    return 1;
  }

  uResult = MsiOpenDatabase(argv[1], MSIDBOPEN_TRANSACT, &target);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiOpenDatabase target"));
    return 1;
  }
  uResult = MsiOpenDatabase(argv[2], MSIDBOPEN_READONLY, &source);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiOpenDatabase source"));
    return 1;
  }
  uResult = MsiDatabaseGenerateTransform(source, target, tempFile, 0, 0);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiDatabaseGenerateTransform"));
    return 1;
  }
  uResult = MsiCreateTransformSummaryInfo(source, target, tempFile, 0, 0);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiCreateTransformSummaryInfo"));
    return 1;
  }
  uResult = MsiDatabaseOpenView(target, TEXT("SELECT `Name`,`Data` FROM _Storages"), &view);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiDatabaseOpenView"));
    return 1;
  }
  record = MsiCreateRecord(2);
  if (!record) {
    std::tcerr << TEXT("error at MsiCreateRecord") << std::endl;
    return 1;
  }
  uResult = MsiRecordSetString(record, 1u, argv[3]);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiRecordSetString substorage-name"));
    return 1;
  }
  uResult = MsiViewExecute(view, record);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiViewExecute"));
    return 1;
  }
  uResult = MsiRecordSetStream(record, 2u, tempFile);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiRecordSetStream substorage-content"));
    return 1;
  }
  uResult = MsiViewModify(view, MSIMODIFY_ASSIGN, record);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiViewModify"));
    return 1;
  }
  uResult = MsiGetSummaryInformation(target, nullptr, 1u, &summaryInfo);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiGetSummaryInformation"));
    return 1;
  }
  uResult = MsiSummaryInfoGetProperty(summaryInfo, PIDSI_TEMPLATE, &dataType, nullptr, nullptr, TEXT(""), &length); // get the length of the stored value
  if (uResult != ERROR_MORE_DATA) {
    HandleMsiError(uResult, TEXT("MsiSummaryInfoGetProperty length"));
    return 1;
  }
  {
    tstring stringValue;
    stringValue.resize(length); // allocate space for the value to be filled (note: reserve would lead to UB!)
    length++; // for the '\0' terminator
    uResult = MsiSummaryInfoGetProperty(summaryInfo, PIDSI_TEMPLATE, &dataType, nullptr, nullptr, stringValue.data(), &length);
    if (uResult != ERROR_SUCCESS) {
      HandleMsiError(uResult, TEXT("MsiSummaryInfoGetProperty value"));
      return 1;
    }
    if (stringValue.find(argv[3]) == tstring::npos) { // check if the required value is already included, only add it if not
      stringValue.append(TEXT(","));
      stringValue.append(argv[3]);
      uResult = MsiSummaryInfoSetProperty(summaryInfo, PIDSI_TEMPLATE, dataType, 0, nullptr, stringValue.data());
      if (uResult != ERROR_SUCCESS) {
        HandleMsiError(uResult, TEXT("MsiSummaryInfoSetProperty"));
        return 1;
      }
    }
  }
  uResult = MsiSummaryInfoPersist(summaryInfo);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiSummaryInfoPersist"));
    return 1;
  }
  uResult = MsiDatabaseCommit(target);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(uResult, TEXT("MsiDatabaseCommit target"));
    return 1;
  }

  bResult = DeleteFile(tempFile);
  if (!bResult) {
    HandleWindowsError(TEXT("DeleteFile temp-file"));
    return 1;
  }

  return 0;
}
