#pragma once
#include <Ole2.h>

// AutoWrap() - Automation helper function...如果使用DISPATCH_METHOD,并且带有参数，需要将参数顺序倒过来！
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...);

// Used to do a simple function mark
void NothingTodo();