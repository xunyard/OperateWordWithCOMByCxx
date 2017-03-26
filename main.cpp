#include <iostream>
#include <stdio.h>
#include "stdafx.h"

#include "MSWord.h"

#define SHOW_PROCESS

void print_info(char *words) {
#ifdef SHOW_PROCESS
	cout << words << endl;
#endif // SHOW_PROCESS
}

using namespace std;

int main() {

	print_info("(0/9)正在启动Word应用程序...");
	MSWord word(false);

	print_info("(1/9)正在打开文件...");
	word.Open(L"D:\\test2_new2.doc");

	print_info("(2/9)插入图片...");
	word.InsertTextAtBookmark(L"PID", L"Test text");
	
	string_table table;
	string_array sa, sb;

	sa.push_back(L"1.1");
	sa.push_back(L"1.2");
	table.push_back(sa);

	sb.push_back(L"2.1");
	sb.push_back(L"2.2");
	table.push_back(sb);

	print_info("(3/9)插入表格...");
	word.InsertTable(word.GoToBookmarkSection(L"PID"), table);

	print_info("(4/9)替换字符串...");
	word.StringFindAndReplace(NULL, L"[名称1]", L"此处为文字替换");

	print_info("(5/9)更新目录...");
	word.UpdateContent();

	print_info("(6/9)修改页眉...");
	word.ReplaceHeaderAndFooter(L"[Type here]", L"修改过后的页眉");

	print_info("(7/9)文件另存为...");
	word.SaveAs(L"D:\\output.doc");

	print_info("(8/9)关闭word...");
	word.Quit();

	print_info("(9/9)完成所有操作，请查看D:\\output.doc文件");
	cout << "Done!";
	return 0;
}