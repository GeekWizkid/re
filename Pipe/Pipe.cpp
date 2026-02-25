// Pipe.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "Pipe.h"
#include "utils.h"
#include <vector>


HANDLE hThread = NULL;
HANDLE hStopEvent = NULL;
SOCKET serverSocket = INVALID_SOCKET;
bool gWsaStarted = false;
std::vector<SOCKET> gClients;
CRITICAL_SECTION gClientsLock;
bool gClientsLockInitialized = false;

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

void EnsureClientLockInitialized()
{
	if(!gClientsLockInitialized) {
		InitializeCriticalSection(&gClientsLock);
		gClientsLockInitialized = true;
	}
}

void AddClientSocket(SOCKET client)
{
	EnsureClientLockInitialized();
	EnterCriticalSection(&gClientsLock);
	gClients.push_back(client);
	LeaveCriticalSection(&gClientsLock);
}

void RemoveClientSocket(SOCKET client)
{
	if(!gClientsLockInitialized)
		return;

	EnterCriticalSection(&gClientsLock);
	for(std::vector<SOCKET>::iterator it = gClients.begin(); it != gClients.end(); ++it) {
		if(*it == client) {
			gClients.erase(it);
			break;
		}
	}
	LeaveCriticalSection(&gClientsLock);
}

void CloseAllClients()
{
	if(!gClientsLockInitialized)
		return;

	EnterCriticalSection(&gClientsLock);
	for(std::vector<SOCKET>::iterator it = gClients.begin(); it != gClients.end(); ++it) {
		if(*it != INVALID_SOCKET)
			closesocket(*it);
	}
	gClients.clear();
	LeaveCriticalSection(&gClientsLock);
}

void BroadcastToClients(const char* message)
{
	if(!message || !gClientsLockInitialized)
		return;

	EnterCriticalSection(&gClientsLock);
	for(std::vector<SOCKET>::iterator it = gClients.begin(); it != gClients.end();) {
		SOCKET client = *it;
		if(send(client, message, (int)strlen(message), 0) == SOCKET_ERROR) {
			closesocket(client);
			it = gClients.erase(it);
			continue;
		}
		++it;
	}
	LeaveCriticalSection(&gClientsLock);
}

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
	if(FAILED(hr) || !pDebugger)
		return 0;

	WSADATA wsaData;
	struct sockaddr_in server, client;
	int c;
	int recvSize;
	char buffer[1024] = {};
	char addrOld[1024] = {};
	char payload[1024] = {};

	if(WSAStartup(MAKEWORD(2,2), &wsaData) == 0)
		gWsaStarted = true;

	if(serverSocket != INVALID_SOCKET)
		closesocket(serverSocket);
	serverSocket = socket(AF_INET, SOCK_STREAM, 0);
	if(serverSocket == INVALID_SOCKET)
		goto Cleanup;

	BOOL reuseAddr = TRUE;
	setsockopt(serverSocket, SOL_SOCKET, SO_REUSEADDR, (const char*)&reuseAddr, sizeof(reuseAddr));

	server.sin_family = AF_INET;
	server.sin_addr.s_addr = INADDR_ANY;
	server.sin_port = htons(8891);
	if(bind(serverSocket, (struct sockaddr *)&server, sizeof(server)) == SOCKET_ERROR)
		goto Cleanup;
	if(listen(serverSocket, SOMAXCONN) == SOCKET_ERROR)
		goto Cleanup;

	c = sizeof(struct sockaddr_in);
	addrOld[0] = 0;
	EnumWindows(EnumWindowsProc, 0);	
	while(WaitForSingleObject(hStopEvent, 0) != WAIT_OBJECT_0) {
		fd_set readSet;
		FD_ZERO(&readSet);
		FD_SET(serverSocket, &readSet);
		SOCKET maxSocket = serverSocket;

		EnsureClientLockInitialized();
		EnterCriticalSection(&gClientsLock);
		for(std::vector<SOCKET>::iterator it = gClients.begin(); it != gClients.end(); ++it) {
			FD_SET(*it, &readSet);
			if(*it > maxSocket)
				maxSocket = *it;
		}
		LeaveCriticalSection(&gClientsLock);

		timeval timeout;
		timeout.tv_sec = 0;
		timeout.tv_usec = 200000;

		int selectResult = select((int)(maxSocket + 1), &readSet, NULL, NULL, &timeout);
		if(selectResult == SOCKET_ERROR)
			break;
		if(selectResult == 0)
			continue;

		if(FD_ISSET(serverSocket, &readSet)) {
			SOCKET acceptedClient = accept(serverSocket, (struct sockaddr *)&client, &c);
			if(acceptedClient != INVALID_SOCKET)
				AddClientSocket(acceptedClient);
		}

		EnterCriticalSection(&gClientsLock);
		for(std::vector<SOCKET>::iterator it = gClients.begin(); it != gClients.end();) {
			SOCKET recvSocket = *it;
			if(!FD_ISSET(recvSocket, &readSet)) {
				++it;
				continue;
			}

			recvSize = recv(recvSocket, payload, sizeof(payload) - 1, 0);
			if(recvSize <= 0) {
				closesocket(recvSocket);
				it = gClients.erase(it);
				continue;
			}

			payload[recvSize] = 0;
			while(recvSize > 0 && (payload[recvSize - 1] == '\r' || payload[recvSize - 1] == '\n')) {
				payload[recvSize - 1] = 0;
				recvSize--;
			}

			lstrcpyA(buffer, "0x");
			lstrcpynA(buffer + 2, payload, sizeof(buffer) - 2);

			if(strcmp(buffer, addrOld)) {
				lstrcpynA(addrOld, buffer, sizeof(addrOld));
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
			++it;
		}
		LeaveCriticalSection(&gClientsLock);
	}

	Cleanup:
	CloseAllClients();
	if(serverSocket != INVALID_SOCKET) {
		closesocket(serverSocket);
		serverSocket = INVALID_SOCKET;
	}
	if(gWsaStarted) {
		WSACleanup();
		gWsaStarted = false;
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
	BroadcastToClients(message);
	return 0;
}

// This is an example of an exported function.
extern "C" PIPE_API void _Connect(HWND hWnd, IDispatch* _pDTE, IDispatch* _pDebugger)
{
	hMain = hWnd;
	pDTE = _pDTE;
	//pDebugger = _pDebugger;
	EnsureClientLockInitialized();

	if(hStopEvent)
		CloseHandle(hStopEvent);
	hStopEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
	if(!hStopEvent)
		return;

	HRESULT hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable, NULL, CLSCTX_INPROC_SERVER, IID_IGlobalInterfaceTable, (void **)&pGIT);
	if(FAILED(hr) || !pGIT)
		return;
	pGIT->RegisterInterfaceInGlobal(_pDebugger, IID_IDispatch, &dwCookie);
	hThread = CreateThread(NULL, 0,	StartServer, NULL, 0, NULL);
}

// This is an example of an exported function.
extern "C" PIPE_API void _Disconnect()
{
	if(hStopEvent)
		SetEvent(hStopEvent);

	if(hThread) {
		WaitForSingleObject(hThread, 3000);
		CloseHandle(hThread);
		hThread = NULL;
	}

	if(hStopEvent) {
		CloseHandle(hStopEvent);
		hStopEvent = NULL;
	}

	if(pGIT) {
		if(dwCookie)
			pGIT->RevokeInterfaceFromGlobal(dwCookie);
		pGIT->Release();
		pGIT = NULL;
		dwCookie = 0;
	}

	CloseAllClients();

	if(serverSocket != INVALID_SOCKET) {
		closesocket(serverSocket);
		serverSocket = INVALID_SOCKET;
	}

	if(gWsaStarted) {
		WSACleanup();
		gWsaStarted = false;
	}

	if(gClientsLockInitialized) {
		DeleteCriticalSection(&gClientsLock);
		gClientsLockInitialized = false;
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
