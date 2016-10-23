// SetOutlook.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#pragma comment(lib,"psapi")
using namespace std;

string OUTLOOK = "OUTLOOK.EXE";
//string OUTLOOK = "calc.exe";

#define CHECK_NULL_RET(bCondition) if (!bCondition) goto Exit0
BOOL EnableDebugPrivilege(void)
{
	HANDLE hToken;
	TOKEN_PRIVILEGES tkp;
	BOOL bRet = FALSE;

	bRet = OpenProcessToken(GetCurrentProcess(),TOKEN_ADJUST_PRIVILEGES|TOKEN_QUERY,&hToken);
	CHECK_NULL_RET(bRet);

	bRet = LookupPrivilegeValue(NULL,SE_DEBUG_NAME,&tkp.Privileges[0].Luid);
	CHECK_NULL_RET(bRet);
	tkp.PrivilegeCount = 1;
	tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
	bRet = AdjustTokenPrivileges(hToken,FALSE,&tkp,0,(PTOKEN_PRIVILEGES)NULL,0);
	CHECK_NULL_RET(bRet);
	bRet = TRUE;

Exit0:
	CloseHandle(hToken);
	return bRet;
}

DWORD FindOutlook()
{
	DWORD aProcesses[1024],cbNeeded,cProcesses;
	unsigned int i;
	if (!EnumProcesses(aProcesses,sizeof(aProcesses),&cbNeeded))
		return 0;
	cProcesses = cbNeeded / sizeof(DWORD);
	HANDLE hProcess;
	DWORD result = 0;
	TCHAR szProcessName[MAX_PATH] = {0};
	string temp;
	for (i = 0; i < cProcesses; ++i)
	{
		hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,FALSE,aProcesses[i]);
		if (hProcess)
		{
			if (GetProcessImageFileName(hProcess,szProcessName,sizeof(szProcessName)/sizeof(TCHAR)))
			{
				temp = szProcessName;
				if (temp.find(OUTLOOK) != -1)
				{
					result = aProcesses[i];
				}
			}
			CloseHandle(hProcess);
		}
		if (result != 0)
			break;
	}
	return result;
}

HMODULE	(WINAPI *RemoteThreadProc)(LPCSTR lpLibFileName);
BOOL DllInject(DWORD dwPid,TCHAR szDllPath[])
{
	if (!EnableDebugPrivilege())
		return FALSE;
	BOOL bRet = FALSE;
	HANDLE hRemoteProcess = OpenProcess(PROCESS_CREATE_THREAD | PROCESS_VM_READ | PROCESS_VM_WRITE |
		PROCESS_VM_OPERATION | PROCESS_QUERY_INFORMATION,FALSE,dwPid);
	RemoteThreadProc = LoadLibraryA;
	DWORD cbSize = (lstrlen(szDllPath) + 1) * sizeof(szDllPath[0]);
	TCHAR *pszRemoteParam = (TCHAR *)VirtualAllocEx(hRemoteProcess,0,cbSize,MEM_COMMIT,PAGE_READWRITE);
	BOOL bWriteMem = WriteProcessMemory(hRemoteProcess,(PVOID)pszRemoteParam,(PVOID)szDllPath,cbSize,NULL);
	HANDLE hThread = CreateRemoteThread(hRemoteProcess,NULL,0,(LPTHREAD_START_ROUTINE)RemoteThreadProc,(LPVOID)pszRemoteParam,0,NULL);
	WaitForSingleObject(hThread,INFINITE);
	DWORD dwExitCode = 0;
	bRet = GetExitCodeThread(hThread,&dwExitCode);
	if (NULL == dwExitCode)
	{
		bRet = FALSE;
	}
	CloseHandle(hThread);
	VirtualFreeEx(hRemoteProcess,(LPVOID)pszRemoteParam,cbSize,MEM_RELEASE);
	CloseHandle(hRemoteProcess);
	return bRet;
}

int _tmain(int argc, _TCHAR* argv[])
{
	DWORD pid = FindOutlook();
	if (pid == 0)
		return -1;
	TCHAR path[MAX_PATH] = {0};
	DWORD length = GetFullPathName("Seed.dll",MAX_PATH,path,NULL);
	if (length == 0)
		return -1;
	DllInject(pid,path);
	return 0;
}

