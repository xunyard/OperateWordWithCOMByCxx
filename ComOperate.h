#pragma once
#include <Ole2.h>

#define SafeRelease(source) NULL == source ? source->Release() : NothingTodo();

class ComOperate {
public:
	ComOperate(IDispatch * Dispatch);

	ComOperate GetSubProperty(WCHAR * Property);

	void SetSimpleProperty(WCHAR * PropertyName, VARIANT &value);

	~ComOperate();

public:
	IDispatch * pDispatch;
};

