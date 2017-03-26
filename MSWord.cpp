#pragma comment(lib, "comsupp.lib")
#include "MSWord.h"
#include "ComOperate.h"
#include <Ole2.h>
#include <comutil.h>


/*
 * === 特别注意 ===
 * 在使用AutoWrap函数时，如果使用DISPATCH_METHOD,并且带有参数，需要将参数顺序倒过来！
 * 再说一遍，顺序倒过来！
 *
 * MSDN上关于word的VBA操作说明：
 * https://msdn.microsoft.com/EN-US/library/office/ff837519.aspx
 * 声明：MFC生成的部分逻辑在这里不能用！不能用！不能用！因为在功能名的生成过程中，它改名字了！
 * 至于为什么MFC能用，因为它根本就没有执行GetIDsOfNames函数，而是直接传入了DISPID值！
 *
 * 通过COM操作Word的过程中，索引号从1开始，而不是一般编程时候的从0开始！
 */


MSWord::MSWord(bool Visable) : 
	COMInitilized(FALSE),
	Application(NULL),
	Documents(NULL),
	Document(NULL)
{

	// Initialize COM for this thread...
	CoInitialize(NULL);

	// Get CLSID for our server...
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Word.Application", &clsid);

	if (FAILED(hr)) {

		::MessageBox(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
		return;
	}

	IDispatch * pApp;
	// Start Word.Application by COM
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pApp);
	Application = new ComOperate(pApp);
	Application->pDispatch;
	

	if (FAILED(hr)) {
		::MessageBox(NULL, "Word not registered properly", "Error", 0x10010);
		return;
	}

	// Make it visible (i.e. app.visible = 1)
	if(Visable)
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		AutoWrap(DISPATCH_PROPERTYPUT, NULL, Application->pDispatch, L"Visible", 1, x);
	}

	// Get Documents collection
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, Application->pDispatch, L"Documents", 0);

		Documents = new ComOperate(result.pdispVal);
	}

	COMInitilized = true;
}


bool MSWord::Open(WCHAR * FileName) {

	if (!COMInitilized) {
		return false;
	}
	else {
		VARIANT result;
		VariantInit(&result);

		AutoWrap(DISPATCH_METHOD, &result, Documents->pDispatch, L"Open", 1, _variant_t(FileName));
		Document = new ComOperate(result.pdispVal);

		return result.pdispVal == NULL;
	}
}

IDispatch * MSWord::GoToBookmarkSection(WCHAR * BookmarkName) {

	IDispatch * pSelection = this->GetSelection();

	VARIANT result;
	VariantInit(&result);

	short GoToBookmark = -1;
	AutoWrap(DISPATCH_METHOD, &result, pSelection, L"GoTo", 4, _variant_t(BookmarkName), vtMissing, vtMissing, _variant_t(GoToBookmark));

	SafeRelease(pSelection);

	return result.pdispVal;
}

void MSWord::InsertText(WCHAR * Text, int FontBold, float FontSize, WCHAR * FontName) {

	IDispatch * pSelection = this->GetSelection(false);

	IDispatch * pFont;
	{
		VARIANT result;
		VariantInit(&result);

		AutoWrap(DISPATCH_PROPERTYGET, &result, pSelection, L"Font", 0);
		pFont = result.pdispVal;
	}

	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pFont, L"Name", 1, _variant_t(FontName));

	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pFont, L"Size", 1, _variant_t(FontSize));

	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pFont, L"Bold", 1, _variant_t(FontBold));

	AutoWrap(DISPATCH_METHOD, NULL, pSelection, L"TypeText", 1, _variant_t(Text));

	SafeRelease(pSelection);
}

void MSWord::InsertTextAtBookmark(WCHAR * BookmarkName, WCHAR * Text, int FontBold, float FontSize, WCHAR * FontName) {
	
	IDispatch * pBookmark = this->GoToBookmarkSection(BookmarkName);
	InsertText(Text, FontBold, FontSize, FontName);
	
	SafeRelease(pBookmark);
}

void MSWord::InsertPicture(WCHAR * BookmarkName, WCHAR * FileName, int index, WCHAR * PictureExplaination) {

	WCHAR * Expression = (WCHAR *)malloc(sizeof(WCHAR) * (lstrlenW(PictureExplaination)+3));

	lstrcpynW(Expression, L"\r\n", 3);
	for (int index = 0; index < lstrlenW(PictureExplaination); index++) Expression[index + 2] = PictureExplaination[index];
	Expression[lstrlenW(PictureExplaination) + 2] = '\0';

	IDispatch *pBookmark = GoToBookmarkSection(BookmarkName);
	this->InsertText(Expression);

	ComOperate InlineShapes = Document->GetSubProperty(L"InlineShapes");

	AutoWrap(DISPATCH_METHOD, NULL, InlineShapes.pDispatch, L"AddPicture", 4, _variant_t(pBookmark), _variant_t(true), _variant_t(false), _variant_t(FileName));

	free(Expression);
	SafeRelease(pBookmark);
}

void MSWord::InsertTable(IDispatch * pRange, string_table &DoubleTable) {

	int RowsCount = DoubleTable.capacity();
	int ColumnsCount = DoubleTable[0].capacity();

	ComOperate Tables = Document->GetSubProperty(L"Tables");

	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_METHOD, &result, Tables.pDispatch, L"Add", 3, _variant_t(ColumnsCount), _variant_t(RowsCount), _variant_t(this->GoToBookmarkSection(L"PID")));
	ComOperate Table(result.pdispVal);

	Table.GetSubProperty(L"Borders").SetSimpleProperty(L"Enable", _variant_t(1));

	ComOperate Range = Table.GetSubProperty(L"Range");
	Range.GetSubProperty(L"Font").SetSimpleProperty(L"Size", _variant_t(10));
	Range.GetSubProperty(L"Font").SetSimpleProperty(L"Name", _variant_t(L"宋体"));
	Range.GetSubProperty(L"ParagraphFormat").SetSimpleProperty(L"Alignment", _variant_t((short)1));//设置段落对齐
	Range.GetSubProperty(L"Cells").SetSimpleProperty(L"VerticalAlignment", _variant_t((short)1));//设置表格元素垂直对齐
	Table.GetSubProperty(L"Rows").SetSimpleProperty(L"Alignment", _variant_t((short)1));//表格居中

	for (int RowIndex = 0; RowIndex < RowsCount; RowIndex++) {
		for (int ColumnIndex = 0; ColumnIndex < ColumnsCount; ColumnIndex++) {

			AutoWrap(DISPATCH_METHOD, &result, Table.pDispatch, L"Cell", 2, _variant_t(ColumnIndex + 1), _variant_t(RowIndex + 1));
			ComOperate Cell(result.pdispVal);

			AutoWrap(DISPATCH_METHOD, NULL, Cell.GetSubProperty(L"Range").pDispatch, L"InsertAfter", 1, _variant_t(DoubleTable[RowIndex][ColumnIndex]));
		}
	}
}

bool MSWord::StringFindAndReplace(IDispatch *pRange, WCHAR * Find, WCHAR * Replace, bool WholeContent) {

	if (true == WholeContent) {
		pRange = Document->GetSubProperty(L"Content").pDispatch;
	}
	
	ComOperate WordFind= ComOperate(pRange).GetSubProperty(L"Find");

	AutoWrap(DISPATCH_METHOD, NULL, WordFind.pDispatch, L"ClearFormatting", 0);

	VARIANT result;
	VariantInit(&result);

	const short wdReplaceAll = -2;

	AutoWrap(DISPATCH_METHOD, &result, WordFind.pDispatch, L"Execute", 11, _variant_t(wdReplaceAll), _variant_t(Replace), vtMissing, vtMissing, vtMissing,
		_variant_t(false), _variant_t(false), vtMissing, _variant_t(true), vtMissing, _variant_t(Find));

	return result.bVal;
}

void MSWord::ReplaceHeaderAndFooter(WCHAR * Find, WCHAR * Replace) {

	ComOperate Sections = Document->GetSubProperty(L"Sections");

	int SectionsCount;
	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, Sections.pDispatch, L"Count", 0);
	SectionsCount = result.intVal;

	for (int index = 1; index < SectionsCount + 1; index++) {

		// Get section
		AutoWrap(DISPATCH_METHOD, &result, Sections.pDispatch, L"Item", 1, _variant_t((long)index));
		
		// Get HeadersFooters
		AutoWrap(DISPATCH_PROPERTYGET, &result, result.pdispVal, L"Headers", 0);

		// Get HeaderFotter
		AutoWrap(DISPATCH_METHOD, &result, result.pdispVal, L"Item", 1, _variant_t(1));

		// Get Range
		AutoWrap(DISPATCH_PROPERTYGET, &result, result.pdispVal, L"Range", 0);

		this->StringFindAndReplace(result.pdispVal, Find, Replace, false);
	}
}

void MSWord::UpdateContent() {

	ComOperate TablesOfContents = Document->GetSubProperty(L"TablesOfContents");

	VARIANT result;
	VariantInit(&result);

	int TableOfContentsCount = 0;
	AutoWrap(DISPATCH_PROPERTYGET, &result, TablesOfContents.pDispatch, L"Count", 0);
	TableOfContentsCount = result.intVal;

	for (int index = 1; index < TableOfContentsCount + 1; index++) {

		AutoWrap(DISPATCH_METHOD, &result, TablesOfContents.pDispatch, L"Item", 1, _variant_t((long)index));

		AutoWrap(DISPATCH_METHOD, NULL, result.pdispVal, L"Update", 0);

		VariantClear(&result);
	}
}

void MSWord::SaveAs(WCHAR * FileName) {
	AutoWrap(DISPATCH_METHOD, NULL, Document->pDispatch, L"SaveAs", 1, _variant_t(FileName));
}

void MSWord::Quit() {

	const short wdDoNotSaveChanges = 0;
	const short wdOriginalDocumentFormat = 1;

	NULL != Document ? AutoWrap(DISPATCH_METHOD, NULL, Document->pDispatch, L"Close", 0) : NothingTodo();
	NULL != Application ? AutoWrap(DISPATCH_METHOD, NULL, Application->pDispatch, L"Quit", 0) : NothingTodo();
}

MSWord::~MSWord() {

	nullptr != Document ? free(Document) : NothingTodo();
	nullptr != Documents ? free(Documents) : NothingTodo();
	nullptr != Application ? free(Application) : NothingTodo();

	// Uninitialize COM for this thread...
	CoUninitialize();
}

IDispatch * MSWord::GetActiveWindow() {

	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, Document->pDispatch, L"ActiveWindow", 0);
	return result.pdispVal;
}

IDispatch * MSWord::GetSelection(bool FromActiveWindow) {

	VARIANT result;
	VariantInit(&result);

	if (FromActiveWindow) {
		IDispatch * pActiveWindow = GetActiveWindow();
		AutoWrap(DISPATCH_PROPERTYGET, &result, pActiveWindow, L"Selection", 0);
		SafeRelease(pActiveWindow);
	}
	else {
		AutoWrap(DISPATCH_PROPERTYGET, &result, Application->pDispatch, L"Selection", 0);
	}

	return result.pdispVal;
}


