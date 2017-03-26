#pragma once

#include <stdio.h>
#include <tchar.h>
#include <vector>
#include <iostream>
#include <Windows.h>

static unsigned int AuthorizationCode = 0xFFFF;


typedef std::vector<double> double_array;
typedef std::vector<double_array> double_table;

typedef std::vector<WCHAR *> string_array;
typedef std::vector<string_array> string_table;

