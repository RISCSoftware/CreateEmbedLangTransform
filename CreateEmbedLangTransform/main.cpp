#include <Windows.h>
#include <tchar.h>
#include <MsiQuery.h>
#include <iostream>

#ifdef _UNICODE
#define tcerr wcerr
#else
#define tcerr cerr
#endif // _UNICODE

void HandleMsiError(LPCTSTR func) {
  PMSIHANDLE error = MsiGetLastErrorRecord();
  if (error) {
    DWORD length = 0;
    UINT status = MsiFormatRecord(0, error, TEXT(""), &length);
    if (status == ERROR_MORE_DATA) {
      length++;
      LPTSTR errorMessage = new TCHAR[length];
      status = MsiFormatRecord(0, error, errorMessage, &length);
      if (status == ERROR_SUCCESS) {
        std::tcerr << TEXT("error at ") << func << TEXT(": ") << errorMessage << std::endl;
      }
      delete[] errorMessage;
    }
  }
}

void HandleWindowsError(LPCTSTR func) {
  DWORD error = GetLastError();
  LPTSTR errorMessage;
  FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, error, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), reinterpret_cast<LPTSTR>(&errorMessage), 0, NULL);
  std::tcerr << TEXT("error at ") << func << TEXT(": ") << errorMessage << std::endl;
  LocalFree(errorMessage);
}

int _tmain(int argc, TCHAR **argv) {
  if (argc < 4) {
    std::tcerr << TEXT("ERROR: too few arguments") << std::endl;
    std::tcerr << TEXT("usage: ") << argv[0] << TEXT(" <target-and-reference-msi> <source-msi> <language-identifier>") << std::endl;
    return 1;
  }

  PMSIHANDLE source, target, view, record;
  UINT uResult;
  DWORD dwResult;
  BOOL bResult;
  TCHAR tempPath[MAX_PATH], tempFile[MAX_PATH];

  dwResult = GetTempPath(MAX_PATH, tempPath);
  if (dwResult > MAX_PATH || dwResult == 0) {
    HandleWindowsError(TEXT("GetTempPath"));
    return 1;
  }
  uResult = GetTempFileName(tempPath, TEXT("CELT"), 0, tempFile);
  if (uResult == 0) {
    HandleWindowsError(TEXT("GetTempFileName"));
    return 1;
  }

  uResult = MsiOpenDatabase(argv[1], MSIDBOPEN_TRANSACT, &target);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiOpenDatabase target"));
    return 1;
  }
  uResult = MsiOpenDatabase(argv[2], MSIDBOPEN_READONLY, &source);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiOpenDatabase source"));
    return 1;
  }
  uResult = MsiDatabaseGenerateTransform(source, target, tempFile, 0, 0);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiDatabaseGenerateTransform"));
    return 1;
  }
  uResult = MsiCreateTransformSummaryInfo(source, target, tempFile, 0, 0);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiCreateTransformSummaryInfo"));
    return 1;
  }
  uResult = MsiDatabaseOpenView(target, TEXT("SELECT `Name`,`Data` FROM _Storages"), &view);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiDatabaseOpenView"));
    return 1;
  }
  record = MsiCreateRecord(2);
  if (!record) {
    std::tcerr << TEXT("error at MsiCreateRecord") << std::endl;
    return 1;
  }
  uResult = MsiRecordSetString(record, 1, argv[3]);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiRecordSetString substorage-name"));
    return 1;
  }
  uResult = MsiViewExecute(view, record);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiViewExecute"));
    return 1;
  }
  uResult = MsiRecordSetStream(record, 2, tempFile);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiRecordSetStream substorage-content"));
    return 1;
  }
  uResult = MsiViewModify(view, MSIMODIFY_ASSIGN, record);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiViewModify"));
    return 1;
  }
  uResult = MsiDatabaseCommit(target);
  if (uResult != ERROR_SUCCESS) {
    HandleMsiError(TEXT("MsiDatabaseCommit target"));
    return 1;
  }

  bResult = DeleteFile(tempFile);
  if (!bResult) {
    HandleWindowsError(TEXT("DeleteFile temp-file"));
    return 1;
  }

  return 0;
}
