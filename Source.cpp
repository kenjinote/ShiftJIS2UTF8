#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "shlwapi")

#include <windows.h>
#include <shlwapi.h>

BOOL CheckCodePage(LPCSTR lpszString, int nCodePage)
{
	int nLength = MultiByteToWideChar(nCodePage, 0, lpszString, -1, 0, 0);
	LPWSTR lpszStringW = (LPWSTR)GlobalAlloc(GPTR, nLength * sizeof(WCHAR));
	MultiByteToWideChar(nCodePage, 0, lpszString, -1, lpszStringW, nLength);
	nLength = WideCharToMultiByte(nCodePage, 0, lpszStringW, -1, 0, 0, 0, 0);
	LPSTR lpszNewStringA = (LPSTR)GlobalAlloc(GPTR, nLength);
	WideCharToMultiByte(nCodePage, 0, lpszStringW, -1, lpszNewStringA, nLength, 0, 0);
	GlobalFree(lpszStringW);
	BOOL bReturnValue = (lstrcmpA(lpszString, lpszNewStringA) == 0);
	GlobalFree(lpszNewStringA);
	return bReturnValue;
}

BOOL ShiftJIS2UTF8(LPCTSTR lpszFilePath, BOOL bBom = TRUE)
{
	if (!lpszFilePath) return FALSE;
	if (!PathFileExists(lpszFilePath)) return FALSE;

	HANDLE hFile = CreateFile(lpszFilePath, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
	if (hFile == INVALID_HANDLE_VALUE) return FALSE;

	const int nFileSize = GetFileSize(hFile, 0);
	if (nFileSize == 0)
	{
		CloseHandle(hFile);
		return FALSE;
	}

	LPSTR lpszText = (LPSTR)GlobalAlloc(GPTR, GetFileSize(hFile, 0) + 1);
	DWORD dwReadSize;
	ReadFile(hFile, lpszText, nFileSize, &dwReadSize, 0);
	CloseHandle(hFile);

	// BOM が付いている場合は既にUTF-8に変換されているとみなす
	if (lpszText[0] == 0xEF && lpszText[0] == 0xBB && lpszText[0] == 0xBF)
	{
		GlobalFree(lpszText);
		return FALSE;
	}

	// 既に UTF-8 の場合
	if (CheckCodePage(lpszText, CP_UTF8))
	{
		GlobalFree(lpszText);
		return FALSE;
	}

	// ANSI コードページではない場合
	if (!CheckCodePage(lpszText, CP_ACP))
	{
		GlobalFree(lpszText);
		return FALSE;
	}

	hFile = CreateFile(lpszFilePath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		GlobalFree(lpszText);
		return FALSE;
	}

	int nLength = MultiByteToWideChar(932, 0, lpszText, -1, 0, 0);
	LPWSTR lpszStringW = (LPWSTR)GlobalAlloc(GPTR, nLength * sizeof(WCHAR));
	MultiByteToWideChar(932, 0, lpszText, -1, lpszStringW, nLength);

	GlobalFree(lpszText);

	nLength = WideCharToMultiByte(CP_UTF8, 0, lpszStringW, -1, 0, 0, 0, 0);
	LPSTR lpszNewStringA = (LPSTR)GlobalAlloc(GPTR, nLength);
	WideCharToMultiByte(CP_UTF8, 0, lpszStringW, -1, lpszNewStringA, nLength, 0, 0);

	GlobalFree(lpszStringW);

	DWORD dwWriteSize;
	if (bBom)
	{
		BYTE byBom[] = { 0xEF, 0xBB, 0xBF };
		WriteFile(hFile, byBom, 3, &dwWriteSize, 0);
	}

	WriteFile(hFile, lpszNewStringA, nLength - 1, &dwWriteSize, 0);

	GlobalFree(lpszNewStringA);

	CloseHandle(hFile);

	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hStatic;
	static HWND hButton;
	switch (msg)
	{
	case WM_CREATE:
		hButton = CreateWindow(TEXT("Button"), TEXT("BOM をつける"), WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hButton, BM_SETCHECK, 1, 0);
		hStatic = CreateWindow(TEXT("Static"), TEXT("ドロップされたShift-JISテキストファイルファイルをUTF-8に変換します。※ファイルは上書きされます。"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_DROPFILES:
		{
			BOOL bBom = (BOOL)SendMessage(hButton, BM_GETCHECK, 0, 0);
			const UINT nFileCount = DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0);
			for (UINT i = 0; i < nFileCount; ++i)
			{
				TCHAR szFilePath[MAX_PATH];
				DragQueryFile((HDROP)wParam, i, szFilePath, _countof(szFilePath));
				ShiftJIS2UTF8(szFilePath, bBom);

			}
			DragFinish((HDROP)wParam);
			MessageBox(hWnd, TEXT("変換が終わりました。"), TEXT("確認"), 0);
		}
		break;
	case WM_SIZE:
		MoveWindow(hStatic, 0, 0, LOWORD(lParam), 32, TRUE);
		MoveWindow(hButton, 0, 32, LOWORD(lParam), 32, TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	TCHAR szClassName[] = TEXT("Window");
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Shift-JIS To UTF-8"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
