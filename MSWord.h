#pragma once
#include "stdafx.h"
#include "COMSupport.h"
#include "ComOperate.h"
using namespace std;

class MSWord {

public:
	MSWord(bool Visable = true);

	// Open Word file, this may take for a while...
	bool Open(WCHAR* FileName);

	IDispatch * GoToBookmarkSection(WCHAR * BookmarkName);

	void InsertText(WCHAR * Text, int FontBold = 0, float FontSize = 12, WCHAR * FontName = L"ËÎÌו");

	void InsertTextAtBookmark(WCHAR * BookmarkName, WCHAR * Text, int FontBold = 0, float FontSize = 12, WCHAR * FontName = L"ËÎÌו");

	void InsertPicture(WCHAR * BookmarkName, WCHAR * FileName, int index, WCHAR * PictureExplaination);

	void InsertTable(IDispatch *pRange, string_table &DoubleTable);

	bool StringFindAndReplace(IDispatch *pRange, WCHAR * Find, WCHAR * Replace, bool WholeContent = true);

	void ReplaceHeaderAndFooter(WCHAR * Find, WCHAR * Replace);

	void UpdateContent();

	void SaveAs(WCHAR * FileName);

	void Quit();

	~MSWord();

private:
	bool COMInitilized;
	ComOperate *Application;
	ComOperate *Documents;
	ComOperate *Document;

private:
	IDispatch * GetActiveWindow();
	IDispatch * GetSelection(bool FromActiveWindow = true);
	
};

