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

const GUID CLSID_Class1 = 
{ 0xC3A7E3DE, 0x8C77, 0x301C, { 0xB5, 0x23, 0x7F, 0x53, 0x78, 0x82, 0x28, 0x69 } };

const GUID IID_Class1 = 
{ 0xC2D5D4CA, 0x0EAC, 0x3671, { 0x9C, 0xE8, 0xE5, 0x1E, 0x18, 0x9D, 0x0B, 0xC7 } };

_IClass1* pIClass1;

int main(int argc, char *argv[])
{
	HRESULT hResult;
	CoInitialize(0);
	hResult = CoCreateInstance(&CLSID_Class1, 0, CLSCTX_INPROC_SERVER, &IID_Class1, (void**)&pIClass1);
	
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

	printf("%S ", (WCHAR *)bstr_out);

	SysFreeString(bstr_in);
	SysFreeString(bstr_out);

	CoUninitialize();
}
