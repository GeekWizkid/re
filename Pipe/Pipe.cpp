// Pipe.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "Pipe.h"
#include "utils.h"


HANDLE hThread;
SOCKET serverSocket, clientSocket;

#ifdef _MANAGED
#pragma managed(push, off)
#endif

struct SearchData
{
	const wchar_t* className;   
	const wchar_t* windowText;   
	HWND foundHwnd;
};

HWND hCodeBrowser;
HWND hMain;
HWND hGoTo;
IDispatch* pDTE;
//IDispatch* pDebugger;
IGlobalInterfaceTable *pGIT;
DWORD dwCookie = 0;

BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
	const int length = GetWindowTextLength(hwnd) + 1;
	wchar_t* buffer = new wchar_t[length];
	GetWindowText(hwnd, buffer, length);

	// ��� ����, ��� �������� ���������� � "CodeBrowser"
	const wchar_t* prefix = L"CodeBrowser:";
	size_t prefixLength = wcslen(prefix);
	if (!wcsncmp(buffer, prefix, prefixLength)) {
		// ����� ����� ��������� �������������� �������� � hwnd
		hCodeBrowser = hwnd;
		delete[] buffer;
		return FALSE; // ���������� �����
	}
	delete[] buffer;
	return TRUE; // ���������� �����
}

void ClickAndTypeText(HWND hwnd, char* text) {
	// ���������� ����� �� ������� ����������
	SetFocus(hwnd);

	// �������� ���������� �������� ����������
	RECT rect;
	GetWindowRect(hwnd, &rect);

	// ��������� ���������� ������ �������� ����������
	POINT pt;
	pt.x = (rect.right + rect.left) / 2;
	pt.y = (rect.bottom + rect.top) / 2;

	// ��������� ���������� � ���������� ���������� �������
	ScreenToClient(hwnd, &pt);

	// �������� ������� ����� ������� ����
	SendMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, MAKELPARAM(pt.x, pt.y));
	SendMessage(hwnd, WM_LBUTTONUP, 0, MAKELPARAM(pt.x, pt.y));

	// ���� ������
	for (size_t i = 0; i < strlen(text); ++i) {
		SendMessage(hwnd, WM_CHAR, text[i], 0);
	}

	// �������� ������� ������� Enter
	PostMessage(hwnd, WM_KEYDOWN, VK_RETURN, 0x001C0001);
	PostMessage(hwnd, WM_KEYUP, VK_RETURN, 0xC01C0001);
}

BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam)
{
	SearchData* pData = reinterpret_cast<SearchData*>(lParam);
	if (!pData) return TRUE; // Продолжать перечисление

	// Проверим класс окна
	wchar_t classBuf[256] = {0};
	if (GetClassName(hwnd, classBuf, 255))
	{
		if (wcscmp(classBuf, pData->className) == 0)
		{
			// Проверим заголовок (caption) окна
			wchar_t textBuf[256] = {0};
			if (GetWindowText(hwnd, textBuf, 255))
			{
				if (wcscmp(textBuf, pData->windowText) == 0)
				{
					// Нашли нужное окно
					pData->foundHwnd = hwnd;
					return FALSE; // Остановить перечисление
				}
			}
		}
	}

	return TRUE;
}

DWORD WINAPI StartServer(LPVOID) 
{	
	IDispatch *pDebugger = NULL;
	HRESULT hr = pGIT->GetInterfaceFromGlobal(dwCookie, IID_IDispatch, (void **)&pDebugger);

	WSADATA wsaData;
	struct sockaddr_in server, client;
	int c;
	int recvSize;
	char buffer[1024] = {};
	char addrOld[1024] = {};
	buffer[0] = '0';
	buffer[1] = 'x'; 

	WSAStartup(MAKEWORD(2,2), &wsaData);
	if(serverSocket) closesocket(serverSocket);
	serverSocket = socket(AF_INET, SOCK_STREAM, 0);
	server.sin_family = AF_INET;
	server.sin_addr.s_addr = INADDR_ANY;
	server.sin_port = htons(8891);
	bind(serverSocket, (struct sockaddr *)&server, sizeof(server));
	listen(serverSocket, 1);
	c = sizeof(struct sockaddr_in);
	clientSocket = accept(serverSocket, (struct sockaddr *)&client, &c);
	EnumWindows(EnumWindowsProc, 0);	
	while(true) {
		SOCKET recvSocket = accept(serverSocket, (struct sockaddr *)&client, &c);
		memcpy(addrOld, buffer, lstrlenA(buffer));
		recvSize = recv(recvSocket, &buffer[2], 18, 0);	
		if(strcmp(buffer, addrOld)) {
			VARIANT result;
			AutoWrap(DISPATCH_PROPERTYGET, &result, pDebugger, L"CurrentMode", 0);
			if(result.lVal == 3) {
				VARIANT variant;
				VariantInit(&variant);

				variant.vt = VT_BOOL;
				variant.boolVal = VARIANT_TRUE;
				AutoWrap(DISPATCH_METHOD, NULL, pDebugger, L"Break", 1, variant);
				VariantClear(&variant);
			}
			VariantClear(&result);

			SearchData data;
			data.className  = L"MsoCommandBar";
			data.windowText = L"Disassembly Window";
			data.foundHwnd  = NULL;

			EnumChildWindows(hMain, EnumChildProc, reinterpret_cast<LPARAM>(&data));


			/*HWND hMDIClient = FindWindowEx(hMain, NULL, L"MDIClient", NULL);
			if (!hMDIClient) goto EndFindWindow;

			HWND hEzMdiContainer = FindWindowEx(hMDIClient, NULL, L"EzMdiContainer", NULL);
			if (!hEzMdiContainer) goto EndFindWindow;

			HWND hDockingView = FindWindowEx(hEzMdiContainer, NULL, L"DockingView", L"Disassembly");
			if (!hDockingView) goto EndFindWindow;

			HWND hGenericPane = FindWindowEx(hDockingView, NULL, L"GenericPane", L"Disassembly");
			if (!hGenericPane) goto EndFindWindow;

			HWND hDisasmPaneWindowClass = FindWindowEx(hGenericPane, NULL, L"DisasmPaneWindowClass", NULL);
			if (!hDisasmPaneWindowClass) goto EndFindWindow;

			HWND hMsoCommandBarDock = FindWindowEx(hDisasmPaneWindowClass, NULL, L"MsoCommandBarDock", L"MsoDockTop");
			if (!hMsoCommandBarDock) goto EndFindWindow;

			HWND hMsoCommandBar = FindWindowEx(hMsoCommandBarDock, NULL, L"MsoCommandBar", L"Disassembly Window");
			if (!hMsoCommandBar) goto EndFindWindow;*/

			HWND hEdit = FindWindowEx(data.foundHwnd, NULL, L"Edit", NULL);
			if (!hEdit) goto EndFindWindow;


			ClickAndTypeText(hEdit, buffer);
		}	
		EndFindWindow:
		closesocket(recvSocket);		
	}
	pDebugger->Release();
	return 0;
}

// ������� ��� �������������� ���� �� ���� ����� �� ����������� ����
void MaximizeWindow(HWND hwnd) {
	// ������������� ���� �� ���� �����	
	if(IsZoomed(hCodeBrowser) == 0) 
	{
		ShowWindow(hwnd, SW_SHOWMAXIMIZED);
		InvalidateRect(hwnd, NULL, TRUE);
		UpdateWindow(hwnd);
	}
	

	// ������������� ����� ��������� ������� � ������ ����, ����� ��� ����� ��������������� �������� ������.
	// ��� ����� ���� �������, ���� ����� ������ ����� ���� � ���������.
	// �������� ������� ������� ������� ������
	//RECT workArea;
	//SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
	int screenWidth = GetSystemMetrics(SM_CXSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYSCREEN);

	// ��������� ���� � ������ ��� ������ ���� ����, �� ��������� ���
	//SetWindowPos(hwnd, HWND_TOP, 0, 0, screenWidth, screenHeight, SWP_NOACTIVATE);
	SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOMOVE);
	SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOMOVE);
}

// This is an example of an exported function.
extern "C" PIPE_API int _WriteAddr(char* message)
{
	DWORD dwWritten = 0;
	if (clientSocket) {
		if (send(clientSocket, message, (int)strlen(message), 0) == SOCKET_ERROR) {
			DWORD dwError = GetLastError();
			if (dwError == 10054) {
				clientSocket = NULL;
				CloseHandle(hThread);
				hThread = CreateThread(NULL, 0,	StartServer, NULL, 0, NULL);
			}
		}
		else {
			//MaximizeWindow(hCodeBrowser);
		}
	}	
	return 0;
}

// This is an example of an exported function.
extern "C" PIPE_API void _Connect(HWND hWnd, IDispatch* _pDTE, IDispatch* _pDebugger)
{
	hMain = hWnd;
	pDTE = _pDTE;
	//pDebugger = _pDebugger;
	HRESULT hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_IGlobalInterfaceTable, (void **)&pGIT);
	pGIT->RegisterInterfaceInGlobal(_pDebugger, IID_IDispatch, &dwCookie);
	hThread = CreateThread(NULL, 0,	StartServer, NULL, 0, NULL);
}

// This is an example of an exported function.
extern "C" PIPE_API void _Disconnect()
{
	if(pGIT)
		pGIT->Release();
	if(hThread)	{
		TerminateThread(hThread, 0);
		CloseHandle(hThread);
		hThread = NULL;
	}
	if(serverSocket) {
		closesocket(serverSocket);
		serverSocket = NULL;
		clientSocket = NULL;
	}	
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		{
			break;
		}
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		_Disconnect();
		break;
	}
    return TRUE;
}

#ifdef _MANAGED
#pragma managed(pop)
#endif


