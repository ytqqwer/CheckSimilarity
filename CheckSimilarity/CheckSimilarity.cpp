﻿// CheckSimilarity.cpp : 定义应用程序的入口点。
//

#include "stdafx.h"
#include "CheckSimilarity.h"

#include <stdio.h>
#include <commdlg.h>
#include <CommCtrl.h>

#include "ExcelReader.h"

#include <ShlObj.h>							//选择文件夹用

#include <codecvt>
#include <io.h>								//遍历文件使用

#define MAX_LOADSTRING 100


const std::wstring PART_OF_SPEECH_DAI = L"r";
const std::wstring PART_OF_SPEECH_DONG = L"v";
const std::wstring PART_OF_SPEECH_FU = L"d";
const std::wstring PART_OF_SPEECH_JIE = L"p";
const std::wstring PART_OF_SPEECH_LIAN = L"c";

const std::wstring PART_OF_SPEECH_LIANG = L"q";
const std::wstring PART_OF_SPEECH_MING = L"n";
const std::wstring PART_OF_SPEECH_NI = L"o";
const std::wstring PART_OF_SPEECH_SHU = L"m";
const std::wstring PART_OF_SPEECH_TAN = L"e";

const std::wstring PART_OF_SPEECH_WEI = L"未";
const std::wstring PART_OF_SPEECH_XING = L"a";
const std::wstring PART_OF_SPEECH_ZHU = L"u";
const std::wstring PART_OF_SPEECH_ZHUI = L"缀";

// 全局变量: 
HINSTANCE hInst;                                // 当前实例
WCHAR szTitle[MAX_LOADSTRING];                  // 标题栏文本
WCHAR szWindowClassMain[MAX_LOADSTRING];            // 主窗口类名

ExcelReader* reader;						//读取器

HWND hWindowMain;							//主窗口句柄

HWND hSearchEdit;							// 搜索框句柄
HWND hPartOfSpeechComboBox;					// 类别下拉列表句柄
HWND hSearchButton;							// 搜索按钮句柄

HWND hDictionaryListView_GKB;				// 词典1列表视图句柄
HWND hDictionaryListView_XH;				// 词典2列表视图句柄

HWND hSimilarityText;						// 相似度
HWND hRelationshipText;						// 对应关系
HWND hNewRelationshipText;					// 新对应关系

HWND hRelationEqualButton;					// 相等按钮
HWND hRelationNotEqualButton;				// 不相等按钮
HWND hRelationUnsureButton;					// 不确定按钮
HWND hRelationBelongButton;					// 属于按钮

HWND hCheckButton;							// 搜索按钮句柄
HWND hNextSenseButton;						// 搜索按钮句柄
HWND hNextWordButton;						// 搜索按钮句柄

											//旧搜索编辑框处理过程
WNDPROC oldEditSearchProc;

// 此代码模块中包含的函数的前向声明: 
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

//搜索编辑框处理过程
LRESULT CALLBACK	subEditSearchProc(HWND, UINT, WPARAM, LPARAM);

#define WTS_U8(wstr) wstringToString(wstr)
std::string wstringToString(const std::wstring& wstr)
{
	LPCWSTR pwszSrc = wstr.c_str();
	int nLen = WideCharToMultiByte(CP_UTF8, 0, pwszSrc, -1, NULL, 0, NULL, NULL);
	if (nLen == 0)
		return std::string("");

	char* pszDst = new char[nLen];
	if (!pszDst)
		return std::string("");

	WideCharToMultiByte(CP_UTF8, 0, pwszSrc, -1, pszDst, nLen, NULL, NULL);
	std::string str(pszDst);
	delete[] pszDst;
	pszDst = NULL;

	return str;
}

#define STW_U8(str) stringToWstring(str)
std::wstring stringToWstring(const std::string& str)
{
	LPCSTR pszSrc = str.c_str();
	int nLen = MultiByteToWideChar(CP_UTF8, 0, pszSrc, -1, NULL, 0);
	if (nLen == 0)
		return std::wstring(L"");

	wchar_t* pwszDst = new wchar_t[nLen];
	if (!pwszDst)
		return std::wstring(L"");

	MultiByteToWideChar(CP_UTF8, 0, pszSrc, -1, pwszDst, nLen);
	std::wstring wstr(pwszDst);
	delete[] pwszDst;
	pwszDst = NULL;

	return wstr;
}

//将TCHAR转为char，*tchar是TCHAR类型指针，*_char是char类型指针   
void TcharToChar(const TCHAR * tchar, char * _char)
{
	int iLength;
	//获取字节长度   
	iLength = WideCharToMultiByte(CP_UTF8, 0, tchar, -1, NULL, 0, NULL, NULL);
	//将tchar值赋给_char    
	WideCharToMultiByte(CP_UTF8, 0, tchar, -1, _char, iLength, NULL, NULL);
}

//把char转为TCHAR
void CharToTchar(const char * _char, TCHAR * tchar)
{
	int iLength;
	iLength = MultiByteToWideChar(CP_UTF8, 0, _char, strlen(_char) + 1, NULL, 0);
	MultiByteToWideChar(CP_UTF8, 0, _char, strlen(_char) + 1, tchar, iLength);
}




int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// 初始化全局字符串
	LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_CHECKSIMILARITY, szWindowClassMain, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// 执行应用程序初始化: 
	if (!InitInstance(hInstance, nCmdShow))
	{
		return FALSE;
	}

	reader = new ExcelReader();

	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_CHECKSIMILARITY));

	MSG msg;
	// 主消息循环: 
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	delete reader;

	return (int)msg.wParam;
}



//
//  函数: MyRegisterClass()
//
//  目的: 注册窗口类。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcexMain;
	wcexMain.cbSize = sizeof(WNDCLASSEX);
	wcexMain.style = CS_HREDRAW | CS_VREDRAW;
	wcexMain.lpfnWndProc = WndProc;
	wcexMain.cbClsExtra = 0;
	wcexMain.cbWndExtra = 0;
	wcexMain.hInstance = hInstance;
	wcexMain.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_CHECKSIMILARITY));
	wcexMain.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wcexMain.hbrBackground = (HBRUSH)(COLOR_WINDOW + 0);
	wcexMain.lpszMenuName = MAKEINTRESOURCEW(IDC_CHECKSIMILARITY);
	wcexMain.lpszClassName = szWindowClassMain;
	wcexMain.hIconSm = LoadIcon(wcexMain.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassExW(&wcexMain);
}

//
//   函数: InitInstance(HINSTANCE, int)
//
//   目的: 保存实例句柄并创建主窗口
//
//   注释: 
//
//        在此函数中，我们在全局变量中保存实例句柄并
//        创建和显示主程序窗口。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance; // 将实例句柄存储在全局变量中

	HWND hWnd = CreateWindowW(szWindowClassMain, szTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, 800, 450, nullptr, nullptr, hInstance, nullptr);

	hWindowMain = hWnd;

	if (!hWnd)
	{
		return FALSE;
	}

	//////////////////////////////////////////////////////////////////////
	//初始化搜索编辑框
	hSearchEdit = CreateWindow(_T("EDIT"), NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_LEFT,
		70, 25, 200, 30, hWnd, (HMENU)ID_SEARCH_EDIT, hInst, NULL);
	oldEditSearchProc = (WNDPROC)SetWindowLongPtr(hSearchEdit, GWLP_WNDPROC, (LONG_PTR)subEditSearchProc);

	//////////////////////////////////////////////////////////////////////
	//初始化类别下拉列表
	hPartOfSpeechComboBox = CreateWindow(WC_COMBOBOX, _T(""), CBS_DROPDOWNLIST | CBS_HASSTRINGS | WS_CHILD | WS_OVERLAPPED | WS_VISIBLE,
		600, 25, 100, 500, hWnd, (HMENU)ID_PART_OF_SPEECH_COMBOBOX, hInst, NULL);

	// load the combobox with item list. Send a CB_ADDSTRING message to load each item
	TCHAR temp[100];

	for (int index = 0; index < 14; index++) {
		LoadStringW(hInstance, ID_PART_OF_SPEECH_N + index, temp, MAX_LOADSTRING);
		SendMessage(hPartOfSpeechComboBox, (UINT)CB_ADDSTRING, (WPARAM)0, (LPARAM)temp);
	}

	// Send the CB_SETCURSEL message to display an initial item in the selection field  
	SendMessage(hPartOfSpeechComboBox, CB_SETCURSEL, (WPARAM)0, (LPARAM)0);

	//////////////////////////////////////////////////////////////////////
	//初始化搜索按钮
	hSearchButton = CreateWindow(_T("BUTTON"), _T("搜索"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		310, 25, 100, 30, hWnd, (HMENU)ID_SEARCH_BUTTON, hInst, NULL);

	//////////////////////////////////////////////////////////////////////
	//初始化文本	
	HWND hWordText = CreateWindow(_T("static"), _T("词语"), WS_CHILD | WS_VISIBLE | SS_LEFT, 30, 30, 30, 30, hWnd,
		NULL, hInst, NULL);

	HWND hPosText = CreateWindow(_T("static"), _T("选择词类"), WS_CHILD | WS_VISIBLE | SS_LEFT, 500, 30, 80, 30, hWnd,
		NULL, hInst, NULL);

	HFONT hFont = CreateFont(20, 0, 0, 0, 0, FALSE, FALSE, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"微软雅黑");//创建字体
	SendMessage(hSearchButton, WM_SETFONT, (WPARAM)hFont, TRUE);//发送设置字体消息
	SendMessage(hSearchEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
	SendMessage(hWordText, WM_SETFONT, (WPARAM)hFont, TRUE);//发送设置字体消息
	SendMessage(hPosText, WM_SETFONT, (WPARAM)hFont, TRUE);

	{
		//////////////////////////////////////////////////////////////////////
		//初始化GKB词典的列表视图
		hDictionaryListView_GKB = CreateWindow(WC_LISTVIEW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_NOSORTHEADER,
			30, 70, 740, 55, hWnd, (HMENU)ID_DICTIONARY_ONE_LISTVIEW, hInst, NULL);

		WCHAR szText[256];     // Temporary buffer.
		int iCol;
		LVCOLUMN lvc;
		// Initialize the LVCOLUMN structure.
		// The mask specifies that the format, width, text,
		// and subitem members of the structure are valid.
		lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;

		// Add the columns.
		for (iCol = 0; iCol < 6; iCol++)
		{
			lvc.iSubItem = iCol;
			lvc.pszText = szText;
			lvc.fmt = LVCFMT_LEFT;		// Left-aligned column.

			if (iCol < 1)				// Width of column in pixels.
				lvc.cx = 70;			//词语
			else if (iCol < 2)
				lvc.cx = 50;			//词类
			else if (iCol < 3)
				lvc.cx = 80;			//拼音
			else if (iCol < 4)
				lvc.cx = 50;			//同形
			else if (iCol < 5)
				lvc.cx = 240;			//释义
			else
				lvc.cx = 250;			//例句

			// Load the names of the column headings from the string resources.
			LoadString(hInst, ID_COLUMN_GKB_WORDS + iCol, szText, sizeof(szText) / sizeof(szText[0]));

			// Insert the columns into the list view.
			if (ListView_InsertColumn(hDictionaryListView_GKB, iCol, &lvc) == -1)
				return FALSE;
		}

		//////////////////////////////////////////////////////////////////////
		//初始化XH词典的列表视图
		hDictionaryListView_XH = CreateWindow(WC_LISTVIEW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_NOSORTHEADER,
			30, 150, 740, 180, hWnd, (HMENU)ID_DICTIONARY_ONE_LISTVIEW, hInst, NULL);

		// Add the columns.
		for (iCol = 0; iCol < 8; iCol++)
		{
			lvc.iSubItem = iCol;
			lvc.pszText = szText;
			lvc.fmt = LVCFMT_LEFT;		// Left-aligned column.

			if (iCol < 1)				// Width of column in pixels.
				lvc.cx = 50;			//ID
			else if (iCol < 2)
				lvc.cx = 60;			//词语
			else if (iCol < 3)
				lvc.cx = 70;			//义项编码
			else if (iCol < 4)
				lvc.cx = 70;			//拼音
			else if (iCol < 5)
				lvc.cx = 50;			//词性
			else if (iCol < 6)
				lvc.cx = 200;			//义项释义
			else if (iCol < 7)
				lvc.cx = 200;			//示例
			else
				lvc.cx = 40;			//相似度

			// Load the names of the column headings from the string resources.
			LoadString(hInst, ID_COLUMN_XH_ID + iCol, szText, sizeof(szText) / sizeof(szText[0]));

			// Insert the columns into the list view.
			if (ListView_InsertColumn(hDictionaryListView_XH, iCol, &lvc) == -1)
				return FALSE;
		}

	}

	//////////////////////////////////////////////////////////////////////
	//Check按钮
	hCheckButton = CreateWindow(_T("BUTTON"), _T("Check"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		670, 350, 100, 30, hWnd, (HMENU)ID_CHECK_BUTTON, hInst, NULL);

	//////////////////////////////////////////////////////////////////////
	//“下一”按钮
	hNextSenseButton = CreateWindow(_T("BUTTON"), _T("下一词义"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		200, 350, 100, 30, hWnd, (HMENU)ID_NEXT_SENSE_BUTTON, hInst, NULL);

	hNextWordButton = CreateWindow(_T("BUTTON"), _T("下一词语"), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		350, 350, 100, 30, hWnd, (HMENU)ID_NEXT_WORD_BUTTON, hInst, NULL);

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	//////////////////////////////////////////////////////////////////////
	// 禁用某些按钮，直到被激活

	return TRUE;
}


// 搜索编辑框 的消息处理函数
LRESULT CALLBACK subEditSearchProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_RETURN:
			//Do your stuff

			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);

			break;  //or return 0; if you don't want to pass it further to def proc
					//If not your key, skip to default:
		}
	default:
		return CallWindowProc(oldEditSearchProc, hWnd, msg, wParam, lParam);
	}
	return 0;
}


void selectColumns()
{

	reader->selectColumn(u8"gkb_词语");
	reader->selectColumn(u8"gkb_词类");
	reader->selectColumn(u8"gkb_拼音");
	reader->selectColumn(u8"gkb_同形");
	reader->selectColumn(u8"gkb_释义");
	reader->selectColumn(u8"gkb_例句");

	reader->selectColumn(u8"ID");
	reader->selectColumn(u8"词语");
	reader->selectColumn(u8"义项编码");
	reader->selectColumn(u8"拼音");
	reader->selectColumn(u8"词性");
	reader->selectColumn(u8"义项释义");
	reader->selectColumn(u8"示例");

	reader->selectColumn(u8"相似度");
}


void resetPartOfSpeech()
{
	// If the user makes a selection from the list:
	//   Send CB_GETCURSEL message to get the index of the selected list item.
	//   Send CB_GETLBTEXT message to get the item.	
	//获取当前列表框中的选项索引
	int ItemIndex = SendMessage(hPartOfSpeechComboBox, (UINT)CB_GETCURSEL, (WPARAM)0, (LPARAM)0);
	//获取当前列表框中的选项值	
	TCHAR  part_of_speech[256];
	(TCHAR)SendMessage(hPartOfSpeechComboBox, (UINT)CB_GETLBTEXT, (WPARAM)ItemIndex, (LPARAM)part_of_speech);

	if (reader->setPartOfSpeech(WTS_U8(part_of_speech))) {
		// 选择列，因为更换词类后，选择的列已被清空
		selectColumns();

	}

}

void setListViewItem(LVITEM& vitem, int iItem, int iSubItem, HWND hWnd, const std::string& data)
{
	std::wstring str = stringToWstring(reader->getCurCellValueInColumn(data));
	vitem.iItem = iItem;
	vitem.iSubItem = iSubItem;
	vitem.pszText = (LPWSTR)str.c_str();

	ListView_InsertItem(hWnd, &vitem);
}

void refreshListView()
{
	// 清除ListView中的所有项 
	ListView_DeleteAllItems(hDictionaryListView_GKB);
	ListView_DeleteAllItems(hDictionaryListView_XH);

	//LVITEM vitem;
	//vitem.mask = LVIF_TEXT;

	/*	先添加项再设置子项内容	*/

	//////////////////////////////////////////////////////////////////////
	//词典1
	//临时变量str保存返回的宽字符字符串	
	LVITEM vitem;
	vitem.mask = LVIF_TEXT;
	vitem.iItem = 0;
	std::wstring str;

	str = stringToWstring(reader->getCurCellValueInColumn(u8"gkb_词语"));
	vitem.iSubItem = 0;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_InsertItem(hDictionaryListView_GKB, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"gkb_词类"));
	vitem.iSubItem = 1;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_GKB, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"gkb_拼音"));
	vitem.iSubItem = 2;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_GKB, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"gkb_同形"));
	vitem.iSubItem = 3;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_GKB, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"gkb_释义"));
	vitem.iSubItem = 4;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_GKB, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"gkb_例句"));
	vitem.iSubItem = 5;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_GKB, &vitem);

	//setListViewItem(vitem,0, 0, hDictionaryListView_GKB, u8"gkb_词语");
	//setListViewItem(vitem,0, 1, hDictionaryListView_GKB, u8"gkb_词类");
	//setListViewItem(vitem,0, 2, hDictionaryListView_GKB, u8"gkb_拼音");
	//setListViewItem(vitem,0, 3, hDictionaryListView_GKB, u8"gkb_同形");
	//setListViewItem(vitem,0, 4, hDictionaryListView_GKB, u8"gkb_释义");
	//setListViewItem(vitem,0, 5, hDictionaryListView_GKB, u8"gkb_例句");

	//////////////////////////////////////////////////////////////////////
	//词典2
	//vitem.iItem = 0;

	str = stringToWstring(reader->getCurCellValueInColumn(u8"ID"));
	vitem.iSubItem = 0;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_InsertItem(hDictionaryListView_XH, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"词语"));
	vitem.iSubItem = 1;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_InsertItem(hDictionaryListView_XH, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"义项编码"));
	vitem.iSubItem = 2;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_XH, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"拼音"));
	vitem.iSubItem = 3;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_XH, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"词性"));
	vitem.iSubItem = 4;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_XH, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"义项释义"));
	vitem.iSubItem = 5;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_XH, &vitem);

	str = stringToWstring(reader->getCurCellValueInColumn(u8"示例"));
	vitem.iSubItem = 6;
	vitem.pszText = (LPWSTR)str.c_str();
	ListView_SetItem(hDictionaryListView_XH, &vitem);

	//setListViewItem(vitem,0, 0, hDictionaryListView_XH, u8"ID");
	//setListViewItem(vitem,0, 1, hDictionaryListView_XH, u8"词语");
	//setListViewItem(vitem,0, 2, hDictionaryListView_XH, u8"义项编码");
	//setListViewItem(vitem,0, 3, hDictionaryListView_XH, u8"拼音");
	//setListViewItem(vitem,0, 4, hDictionaryListView_XH, u8"词性");
	//setListViewItem(vitem,0, 5, hDictionaryListView_XH, u8"义项释义");
	//setListViewItem(vitem,0, 6, hDictionaryListView_XH, u8"示例");

}


void refreshMainWindow()
{

	refreshListView();

	//////////////////////////////////////////////////////////////////////
	//TODO 设置搜索框中的文本

	//str = stringToWstring(reader->getCurCellValueInColumn(u8"相似度"));
	//SetWindowText(hSimilarityText, (LPWSTR)str.c_str());		// 相似度

	//str = stringToWstring(reader->getCurCellValueInColumn(u8"映射关系"));
	//SetWindowText(hRelationshipText, (LPWSTR)str.c_str());		// 对应关系

}


// 激活某些控件
void activeControls()
{
	//TODO

}

//
//  函数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目的:    处理主窗口的消息。
//
//  WM_COMMAND  - 处理应用程序菜单
//  WM_PAINT    - 绘制主窗口
//  WM_DESTROY  - 发送退出消息并返回
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_COMMAND:
	{
		// 分析wParam全部: 
		//switch (wParam)
		//{
		//case CBN_SELCHANGE:
		//{
		//
		//}
		//break;
		//default:
		//	//继续处理
		//	break;
		//}

		// 分析wParam高位: 
		switch (HIWORD(wParam))
		{
		case CBN_SELCHANGE:
		{
			//只有当读取器已经加载过文件时，才自动刷新主窗口
			//此时才需要替换词类，打开文件时也会根据列表内容再选择一次词类
			if (reader->isExistingFile()) {

				resetPartOfSpeech();
				refreshMainWindow();
			}

		}
		break;
		default:
			//继续处理
			break;
		}

		// 分析wParam低位: 
		switch (LOWORD(wParam))
		{
		case ID_SEARCH_BUTTON:
		{

			MessageBox(hWnd, L"您点击了一个按钮。", L"提示", MB_OK);

		}
		break;
		case ID_CHECK_BUTTON:
			MessageBox(hWnd, L"点了个按钮。", L"提示", MB_OK);
			break;
		case ID_NEXT_WORD_BUTTON:
		{
			if (reader->isExistingFile()) {
				if (reader->nextWord()) {
					refreshMainWindow();
				}
				else
					MessageBox(hWnd, L"已是最后一个词语。", L"提示", MB_OK);
			}
		}
		break;
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		case IDM_OPEN:
		{
			//调用 shell32.dll api   调用浏览文件夹对话框 
			TCHAR szPathName[MAX_PATH];
			BROWSEINFO bInfo = { 0 };
			bInfo.hwndOwner = GetForegroundWindow();//父窗口  
			bInfo.lpszTitle = TEXT("请选择词义对应表所在的文件夹");
			bInfo.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI/*包含一个编辑框 用户可以手动填写路径 对话框可以调整大小之类的..*/ |
				BIF_UAHINT/*带TIPS提示*/ | BIF_NONEWFOLDERBUTTON /*不带新建文件夹按钮*/;
			LPITEMIDLIST lpDlist;
			lpDlist = SHBrowseForFolder(&bInfo);
			if (lpDlist != NULL)//单击了确定按钮  
			{
				//首先清除之前载入的所有信息
				reader->clear();

				////////////////////////////////////////////////////////////////////
				// 获取目录，为了支持中文，需要将获取的字符串的TCHAR编码转为CHAR编码，使用代码页CP_ACP
				SHGetPathFromIDList(lpDlist, szPathName);

				char tempPathName[256];
				int iLength;
				//获取字节长度   
				iLength = WideCharToMultiByte(CP_ACP, 0, szPathName, -1, NULL, 0, NULL, NULL);
				//将tchar值赋给_char    
				WideCharToMultiByte(CP_ACP, 0, szPathName, -1, tempPathName, iLength, NULL, NULL);

				////////////////////////////////////////////////////////////////////
				//搜索目录下的xlsx文件
				std::string searchPath = tempPathName;
				searchPath = searchPath + "\\*.xlsx";
				_finddata_t fileDir;
				long lfDir;
				if ((lfDir = _findfirst(searchPath.c_str(), &fileDir)) == -1l) {
					MessageBox(hWindowMain, L"No xlsx file is found\n", L"提示", MB_OK);
					break;
				}
				else {
					do {
						//找到了xlsx文件，先将char转换为wchar，再将wchar转换为utf-8编码的char
						TCHAR temp1[256];
						int iLength;
						iLength = MultiByteToWideChar(CP_ACP, 0, fileDir.name, strlen(fileDir.name) + 1, NULL, 0);
						MultiByteToWideChar(CP_ACP, 0, fileDir.name, strlen(fileDir.name) + 1, temp1, iLength);

						std::wstring wtemp1(temp1);
						std::string utf8_str = WTS_U8(wtemp1);

						reader->addXlsxFileName(utf8_str);

					} while (_findnext(lfDir, &fileDir) == 0);
				}
				_findclose(lfDir);

				////////////////////////////////////////////////////////////////////
				//传入正则表达式，以及代表的词类英文字母
				std::wstring pattern_1to1{ L"\\w*\\dto\\d\\w*" };
				std::wstring pattern_1tom{ L"\\w*\\dto[a-zA-Z]\\w*" };
				std::wstring pattern_mto1{ L"\\w*[a-zA-Z]to\\d\\w*" };
				std::wstring pattern_mtom{ L"\\w*[a-zA-Z]to[a-zA-Z]\\w*" };
				std::wstring pattern_end{ L"\\w*.xlsx" };

				std::vector<std::wstring> pattern_part;
				pattern_part.push_back(pattern_1to1);
				pattern_part.push_back(pattern_1tom);
				pattern_part.push_back(pattern_mto1);
				pattern_part.push_back(pattern_mtom);

				std::wstring wpath = szPathName;
				wpath = wpath + L"\\";

				for (auto& part : pattern_part) {
					reader->loadXlsxFile(WTS_U8(part + L"代" + pattern_end), WTS_U8(PART_OF_SPEECH_DAI), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"动" + pattern_end), WTS_U8(PART_OF_SPEECH_DONG), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"副" + pattern_end), WTS_U8(PART_OF_SPEECH_FU), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"介" + pattern_end), WTS_U8(PART_OF_SPEECH_JIE), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"连" + pattern_end), WTS_U8(PART_OF_SPEECH_LIAN), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"量" + pattern_end), WTS_U8(PART_OF_SPEECH_LIANG), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"名" + pattern_end), WTS_U8(PART_OF_SPEECH_MING), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"拟" + pattern_end), WTS_U8(PART_OF_SPEECH_NI), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"数" + pattern_end), WTS_U8(PART_OF_SPEECH_SHU), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"叹" + pattern_end), WTS_U8(PART_OF_SPEECH_TAN), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"未" + pattern_end), WTS_U8(PART_OF_SPEECH_WEI), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"形" + pattern_end), WTS_U8(PART_OF_SPEECH_XING), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"助" + pattern_end), WTS_U8(PART_OF_SPEECH_ZHU), WTS_U8(wpath));
					reader->loadXlsxFile(WTS_U8(part + L"缀" + pattern_end), WTS_U8(PART_OF_SPEECH_ZHUI), WTS_U8(wpath));
				}

				////////////////////////////////////////////////////////////////////
				// 打开文件后默认显示当前词类选择列表中指定的表格，刷新主窗口
				// 这样当未打开文件时，也可以随便选择词类，但是不会产生效果
				resetPartOfSpeech();

				//////////////////////////////////////////////////////////////////////
				//TODO 激活某些控件
				activeControls();


				refreshMainWindow();

			}

		}
		break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}

		//// 分析lParam全部: 
		//switch (lParam)
		//{
		//case ID_SEARCH_BUTTON:
		//{
		//	MessageBox(hWnd, L"您点击了一个按钮。", L"提示", MB_OK);
		//}
		//break;
		//default:
		//	return DefWindowProc(hWnd, message, wParam, lParam);
		//}


	}
	break;
	case WM_PAINT:
	{
		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(hWnd, &ps);
		// TODO: 在此处添加使用 hdc 的任何绘图代码...

		EndPaint(hWnd, &ps);
	}
	break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}

	return 0;
}

// “关于”框的消息处理程序。
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
