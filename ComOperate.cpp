#include "ComOperate.h"
#include "COMSupport.h"


ComOperate::ComOperate(IDispatch * Dispatch) : pDispatch(Dispatch) {

}

ComOperate ComOperate::GetSubProperty(WCHAR * Property) {

	if (nullptr == pDispatch) {
		return nullptr;
	}

	VARIANT result;
	VariantInit(&result);
	AutoWrap(DISPATCH_PROPERTYGET, &result, pDispatch, Property, 0);

	return ComOperate(result.pdispVal);
}

void ComOperate::SetSimpleProperty(WCHAR * PropertyName, VARIANT & value) {

	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pDispatch, PropertyName, 1, value);
}

ComOperate::~ComOperate() {
	SafeRelease(pDispatch);
}
