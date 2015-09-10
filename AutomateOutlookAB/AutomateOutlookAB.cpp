// AutomateOutlookAB.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "AutomateOutlookAB.h"

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);
BOOL CALLBACK		EnumChildWindowProc(HWND hwnd, LPARAM lParam);

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPTSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_AUTOMATEOUTLOOKAB, szWindowClass, MAX_LOADSTRING);
	/*MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}
	*/
	/*
	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_AUTOMATEOUTLOOKAB));

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	*/

	HWND MainWindowHandle = FindWindow(0, L"Add Users");

	std::vector<HWND> ChildList;

	EnumChildWindows(MainWindowHandle, EnumChildWindowProc, (LPARAM)&ChildList);

	std::ifstream MemberListFile;
	std::string line;
	MemberListFile.open("c:\\users\\1242590410N\\Desktop\\input.txt");

	//SendMessage(ChildList[27], CB_SELECTSTRING, -1, (LPARAM)L"All Groups");
	SetForegroundWindow(MainWindowHandle);
	while (std::getline( MemberListFile, line))
	{
		char* Message = "";
		SendMessageA(ChildList[4], WM_SETTEXT, strlen(Message), (WPARAM)Message);
		for (int i = 0; i < line.length(); i++)
		{
			SendMessage(ChildList[4], WM_CHAR, line[i], 1);
		}
		bool SelectionMade = false;
		int Selections[5] = { 0, 0, 0, 0, 0 };
		
		//Sleep(2000);
			/*
		while (!SelectionMade)
		{
			if (SendMessage(ChildList[10], LB_GETSELITEMS, 1, (LPARAM)Selections) != LB_ERR)
			{
				WCHAR buffer[256];
				SendMessage(ChildList[10], LB_GETTEXT, Selections[0], (LPARAM)buffer);
				MessageBox(0, buffer, L"Result!", 0);
			}
			else
			{
				MessageBox(0, L"SingleSelect", L"Error", 0);
			}
		}
		*/
		SendMessage(ChildList[13], BM_CLICK, 0, 0);
	}

	MemberListFile.close();
	/*char* DLName = "USAF JB A-NAFW NGB A3 List FAL";
	for (int i = 0; i < strlen(DLName); i++)
	{
		SendMessage(ChildList[4], WM_CHAR, DLName[i], 1);
	}
	*/
	return 0;
}

BOOL CALLBACK EnumChildWindowProc(HWND hwnd, LPARAM lParam)
{
	std::vector<HWND>* List = (std::vector<HWND>*)lParam;
	List->push_back(hwnd);
	return true;
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_AUTOMATEOUTLOOKAB));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_AUTOMATEOUTLOOKAB);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;

   hInst = hInstance; // Store instance handle in our global variable

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message)
	{
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
