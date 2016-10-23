// Seed.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
HMODULE hMod;
HWND hWindow = NULL;
DWORD tid = 0;

#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lp)                        ((int)(short)LOWORD(lp))
#endif
#ifndef GET_Y_LPARAM
#define GET_Y_LPARAM(lp)                        ((int)(short)HIWORD(lp))
#endif

LRESULT CALLBACK SuppressClose(int nCode,WPARAM wParam,LPARAM lParam)
{
	switch (nCode)
	{
	case HCBT_SYSCOMMAND:
		if (wParam == SC_CLOSE && hWindow && tid != 0)
		{
			POINT p;
			p.x = GET_X_LPARAM(lParam);
			p.y = GET_Y_LPARAM(lParam);
			HWND hwnd = WindowFromPoint(p);
			if (hwnd == NULL)
				break;
			char text[256] = {0};
			if (GetWindowText(hwnd,text,sizeof(text)) == 0)
				break;
			if (strstr(text,"Microsoft Outlook"))
			{
				PostMessage(hwnd,WM_SYSCOMMAND,SC_MINIMIZE,0);
				return TRUE;
			}
		}
		break;
	default:
		break;
	}
	return CallNextHookEx(NULL,nCode,wParam,lParam);
}

void ShowError(const char *msg)
{
	char errmsg[64] = {0};
	sprintf_s(errmsg,sizeof errmsg,"%s Error:%d",msg,GetLastError());
	MessageBox(NULL,errmsg,0,0);
}


BOOL CALLBACK EnumWindowsProc(HWND hwnd,LPARAM lParam)
{
	DWORD processID;
	DWORD *result = (DWORD *)lParam;
	DWORD threadID = GetWindowThreadProcessId(hwnd,&processID);
	if (processID == GetCurrentProcessId())
	{
		hWindow = hwnd;
		*result = threadID;
		return FALSE;
	}
	else
		return TRUE;
}

DWORD __stdcall MainProc(void *param)
{
	DWORD threadID = 0;
	EnumWindows(EnumWindowsProc,(LPARAM)&threadID);
	if (threadID == 0)
	{
		ShowError("EnumWindows");
		return 0;
	}

	tid = threadID;
	HHOOK hook = SetWindowsHookEx(WH_CBT,SuppressClose,NULL,threadID);
	if (!hook)
	{
		ShowError("SetWindowsHookEx");
		return 0;
	}

	MessageBox(NULL,"…Ë÷√≥…π¶","Seed",MB_ICONINFORMATION);
	MSG msg;
	while (GetMessage(&msg,NULL,0,0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	UnhookWindowsHookEx(hook);
	return 0;
}

void Init(HMODULE hModule)
{	
	hMod = hModule;
	CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)MainProc,NULL,0,NULL);
}	

