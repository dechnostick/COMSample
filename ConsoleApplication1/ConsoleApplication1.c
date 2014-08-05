#include <stdio.h>
#include "objbase.h"

typedef struct tagMyVTable {
	void *QueryInterface;
	void *AddRef;
	ULONG(__stdcall *Release)(void* _this);
	void *GetTypeInfoCount;
	void *GetTypeInfo;
	void *GetIdsOfNames;
	void *Invoke;
	void *ToString;
	void *Equals;
	void *GetHashCode;
	void *GetType;
	HRESULT(__stdcall *Test)(void* _this, BSTR str, BSTR* pRetValue);
} MyVTable;

typedef struct tagIClass1 {
	MyVTable *lpVtbl;
} _IClass1;

CLSID clsid;
_IClass1* pIClass1;

int main(int argc, char *argv[])
{
	HRESULT hResult;

	OleInitialize(NULL);
	CLSIDFromProgID(L"CommonVB.SpreadsheetGearWrapper", &clsid);

	CoInitialize(0);
	hResult = CoCreateInstance(&clsid, 0, CLSCTX_INPROC_SERVER, &IID_IUnknown, (void**)&pIClass1);

	if (!SUCCEEDED(hResult)) {
		printf("ERROR: %08x\n", hResult);
		return 0;
	}

	char in[128];
	char out[128];

	OLECHAR ole_in[256];
	OLECHAR ole_out[256];

	BSTR bstr_in = NULL;
	BSTR bstr_out = NULL;

	memset(in, 0x00, sizeof(in));
	memset(out, 0x00, sizeof(out));

	strcat(in, "world");

	mbstowcs(ole_in, in, sizeof(in));
	mbstowcs(ole_out, out, sizeof(out));

	bstr_in = SysAllocString(ole_in);
	bstr_out = SysAllocString(ole_out);

	pIClass1->lpVtbl->Test((void*)pIClass1, bstr_in, &bstr_out);

	printf("%S\n", (WCHAR *)bstr_out);

	SysFreeString(bstr_in);
	SysFreeString(bstr_out);

	CoUninitialize();
}
