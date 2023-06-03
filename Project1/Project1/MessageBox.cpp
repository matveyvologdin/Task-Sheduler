#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>

int _stdcall WinMain(HINSTANCE hInstance, HINSTANCE NotUsed, LPSTR pCmdLine, int fShow)
{
	FreeConsole();
	if (strlen(pCmdLine) > 5)
	{
		char message[MAX_PATH] = "Blocked ICMP from IP: ";
		strcat(message, pCmdLine);
		MessageBoxA(NULL, message, "Notification", MB_OK);
	}
	else
	{
		MessageBoxA(NULL, "Winsows security is changed!", "Notification", MB_OK);
	}
	return EXIT_SUCCESS;
}