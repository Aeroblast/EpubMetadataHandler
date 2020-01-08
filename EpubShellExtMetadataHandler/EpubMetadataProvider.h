#pragma once
#include <windows.h>
#include <propsys.h>
#include<propkey.h>
#include<propvarutil.h>
#include "unzip.h"
#pragma comment(lib,"Propsys.lib")

class EPUB_METADATA
{
public:
	CABSTR creator;
	LPWSTR title=0,lang=0;
	EPUB_METADATA()
	{
		creator.cElems= 1;
		creator.pElems= (LPWSTR*)CoTaskMemAlloc(creator.cElems * sizeof(LPWSTR));
		if(creator.pElems)memset(creator.pElems,0, creator.cElems * sizeof(LPWSTR));
	}
	~EPUB_METADATA()
	{
		
		for (DWORD i = 0; i < creator.cElems; i++)
		{
			if(creator.pElems[i]!=0)SysFreeString(creator.pElems[i]);
		}
		CoTaskMemFree(creator.pElems);
		if (title)SysFreeString(title);
		if (lang)SysFreeString(lang);
		
	}
	BSTR AllocStr(char* s, int length)
	{
		int wcsLen = MultiByteToWideChar(CP_UTF8, NULL, s, length, NULL, 0);
		BSTR r = SysAllocStringLen(NULL, (UINT)wcslen + 1);
		if (r) 
		{
			MultiByteToWideChar(CP_UTF8, NULL, s, length, r, wcsLen);
			r[wcsLen] = 0;
		}
		return r;
	}
	void AddCreator(char* s, int length)
	{
		*(creator.pElems) = AllocStr(s,length);
	}
	void SetTitle(char* s, int length) 
	{
		title = AllocStr(s, length);
	}
	void SetLang(char* s, int length) 
	{
		lang = AllocStr(s, length);
	}
};

class EpubMetadataProvider : 
    public IInitializeWithStream, 
    public IPropertyStore
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode);



    // Meta
	HRESULT GetCount(
		DWORD *cProps
	);
	//S_OK
	HRESULT GetAt(
		DWORD       iProp,
		PROPERTYKEY *pkey
	);
	HRESULT GetValue(
		REFPROPERTYKEY key,
		PROPVARIANT    *pv
	);
	HRESULT Commit();
	HRESULT SetValue(
		REFPROPERTYKEY key,
		REFPROPVARIANT propvar
	);
 

    EpubMetadataProvider();
	PROPERTYKEY keys[3] = { INIT_PKEY_Language, INIT_PKEY_Title, INIT_PKEY_Author };
protected:
    ~EpubMetadataProvider();

private:
    // Reference count of component.
    long m_cRef;

    // Provided during initialization.
    IStream *m_pStream;

	
	EPUB_METADATA meta;

};


