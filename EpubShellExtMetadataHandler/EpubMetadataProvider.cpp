/********************************  ********************************\
 修改自RecipeThumbnailProvider.cpp

\*******************************************************************************/

#include "EpubMetadataProvider.h"
#include <Shlwapi.h>
#include <Wincrypt.h>   // For CryptStringToBinary.
#pragma comment(lib, "Shlwapi.lib")
#include<regex>
//debug
#include<string>
#include<fstream>
#include <iostream>
#ifdef _DEBUG
#define DEBUGLOG(str) {std::wofstream file;file.open("E:\\log.txt", std::ios::app);file << str<< "\n";}
#else
#define DEBUGLOG(str) ;
#endif // DEBUG


//helpers
HRESULT OPFHandle(char* text, EPUB_METADATA& meta);
HRESULT  GetOPFPath(char* text, BSTR* opf_path);
HRESULT GetMetadata(HZIP, EPUB_METADATA&);



extern HINSTANCE g_hInst;
extern long g_cDllRef;


EpubMetadataProvider::EpubMetadataProvider() : m_cRef(1), m_pStream(NULL)
{
	DEBUGLOG("Handle");
	InterlockedIncrement(&g_cDllRef);
}


EpubMetadataProvider::~EpubMetadataProvider()
{
	DEBUGLOG("~Handle");
	InterlockedDecrement(&g_cDllRef);
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP EpubMetadataProvider::QueryInterface(REFIID riid, void** ppv)
{
	static const QITAB qit[] =
	{
		QITABENT(EpubMetadataProvider,IPropertyStore),
		QITABENT(EpubMetadataProvider, IInitializeWithStream),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) EpubMetadataProvider::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) EpubMetadataProvider::Release()
{
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (0 == cRef)
	{
		delete this;
	}

	return cRef;
}

#pragma endregion


#pragma region IInitializeWithStream

// Initializes the thumbnail handler with a stream.
IFACEMETHODIMP EpubMetadataProvider::Initialize(IStream* pStream, DWORD grfMode)
{
	// A handler instance should be initialized only once in its lifetime. 
	HRESULT hr = HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);

	if (m_pStream == NULL)
	{
		// Take a reference to the stream if it has not been initialized yet.
		hr = pStream->QueryInterface(&m_pStream);
		if (!SUCCEEDED(hr)) { return E_FAIL; }
		//
		ULONG actRead;
		ULARGE_INTEGER fileSize;
		DEBUGLOG("run");
		hr = IStream_Size(m_pStream, &fileSize); if (!SUCCEEDED(hr)) { DEBUGLOG("Error at IStream_Size"); return E_FAIL; }

		byte* buffer = (byte*)CoTaskMemAlloc(fileSize.QuadPart);
		if (!buffer) { DEBUGLOG("Error at malloc buffer");  return E_FAIL; }
		else
		{
			hr = m_pStream->Read(buffer, (LONG)fileSize.QuadPart , &actRead);
			if (!SUCCEEDED(hr))
			{
				CoTaskMemFree(buffer);
				DEBUGLOG("Error at Read Stream");
				return E_FAIL;
			}
		}
		HZIP zip = OpenZip(buffer, actRead, 0);
		if (!zip) { DEBUGLOG("Error at Load zip"); CoTaskMemFree(buffer); return E_FAIL; }

		GetMetadata(zip,meta);
		CoTaskMemFree(buffer);
		CloseZip(zip);
		DEBUGLOG("End Read");
	}
	return hr;
}

#pragma endregion


#pragma region IPropertyStore
// Meta
HRESULT EpubMetadataProvider::GetCount(DWORD* cProps) {
	*cProps = 3;
	return S_OK;
}
//S_OK
HRESULT EpubMetadataProvider::GetAt(
	DWORD       iProp,
	PROPERTYKEY* pkey
) 
{
	*pkey = keys[iProp];
	return S_OK;
}

//return S_OK or INPLACE_S_TRUNCATED if successful, or an error value otherwise。
//INPLACE_S_TRUNCATED is returned to indicate that
//the returned PROPVARIANT was converted into a more canonical form. 
//For example, this would be done to trim leading or trailing spaces from a string value. 

//If the PROPERTYKEY referenced in key is not present in the property store, 
//this method returns S_OK and the vt member of the structure that is pointed to by pv is set to VT_EMPTY.
HRESULT EpubMetadataProvider::GetValue(
	REFPROPERTYKEY key,
	PROPVARIANT* pv
) 
{
	PropVariantInit(pv);
	if (IsEqualPropertyKey(key, keys[0]))
	{
		InitPropVariantFromString(meta.lang,pv);
		DEBUGLOG("lang");
	}
	else if (IsEqualPropertyKey(key, keys[1]))
	{
		InitPropVariantFromString(meta.title, pv);
		DEBUGLOG("title");
	}
	else if (IsEqualPropertyKey(key, keys[2])) 
	{
		//InitPropVariantFromStringVector(meta.creator.pElems,meta.creator.cElems,pv);
		InitPropVariantFromStringAsVector(meta.creator.pElems[0],pv);
		DEBUGLOG("creator");
	}
	else return E_FAIL;
	return INPLACE_S_TRUNCATED;
}
HRESULT EpubMetadataProvider::Commit() 
{
	return STG_E_ACCESSDENIED;//This is an error code. The property store was read-only so the method was not able to set the value.
}
HRESULT EpubMetadataProvider::SetValue(REFPROPERTYKEY key,REFPROPVARIANT propvar)
{
	return STG_E_ACCESSDENIED;
}
#pragma endregion


#pragma region Helper Functions



HRESULT GetMetadata(HZIP zip, EPUB_METADATA&meta)
{

	int index = -1; ZIPENTRY entry;
	HRESULT hr = FindZipItem(zip, L"META-INF/container.xml", false, &index, &entry);
	if (index == -1) { return E_FAIL; }
	char* xml = (char*)CoTaskMemAlloc(entry.unc_size);
	if (xml)
	{
		hr = UnzipItem(zip, index, xml, entry.unc_size);
		if (SUCCEEDED(hr)) //获得container.xml
		{
			DEBUGLOG("Got container.xml");

			BSTR opf_path = NULL;
			hr = GetOPFPath(xml, &opf_path);//解析container.xml
			if (SUCCEEDED(hr))//得到opf路径
			{
				DEBUGLOG(opf_path);
				hr = FindZipItem(zip, opf_path, false, &index, &entry);
				if (index != -1)
				{
					char* opf_data = (char*)CoTaskMemAlloc(entry.unc_size);
					if (opf_data)
					{
						hr = UnzipItem(zip, index, opf_data, entry.unc_size);
						if (SUCCEEDED(hr))
						{
							DEBUGLOG("Got opf data");
							hr = OPFHandle(opf_data, meta);
						}
						CoTaskMemFree(opf_data);
					}
					else { hr = E_FAIL; }
				}
				else
				{
					hr = E_FAIL;
				}
				SysFreeString(opf_path);
			}
		}
		CoTaskMemFree(xml);
	}

	return hr;
}

HRESULT  GetOPFPath(char* text, BSTR* opf_path)
{
	std::regex rx("full-path=\"(.*?)\"");
	std::string t = text;
	std::smatch sm;
	std::regex_search(t, sm, rx);
	if (sm.length() > 1)
	{
		std::string r = sm[1].str();
		const size_t strsize = r.length() + 1;
		*opf_path = SysAllocStringLen(NULL, (UINT)strsize);
		mbstowcs(*opf_path, r.c_str(), strsize);
		return S_OK;
	}
	else
	{
		return E_FAIL;
	}
}
bool CompareTag(const char*a,char*b, int b_i,int b_l) 
{
	int length =(int) strlen(a);
	if (b_l != length)return false;
	for (int i = 0; i < length; i++)if (a[ i] != b[b_i + i])return false;
	return true;
}
struct {
	int start;//pos of <
	int end;
	int tagname_length;
	int inner_start;
	int inner_length;
}elementInfo;
bool Element(char* text,int start) 
{
	elementInfo.start = start;
	if (text[start + 1] == '!'|| text[start + 1] == '?'|| text[start + 1] == '/')return false;
	int pos = start + 1;
	while (text[pos] != ' '&&text[pos] != '>')pos++;
	elementInfo.tagname_length = pos - start-1;
	while (text[pos] != '>')pos++;
	elementInfo.inner_start = pos + 1;
	while (text[pos] != '<')pos++;
	elementInfo.inner_length = pos-elementInfo.inner_start;
	while (text[pos] != '>')pos++;
	elementInfo.end =pos;
	return true;
}
HRESULT OPFHandle(char* text, EPUB_METADATA& meta)
{
	const char* _tagname_metadata = "metadata";
	const char* _tagname_creater = "dc:creator";
	const char* _tagname_title = "dc:title";
	const char* _tagname_lang= "dc:language";
	int i = 0;
#define STATE_NORMAL 0
#define STATE_IN_METADATA 1
	int state = STATE_NORMAL;
	void (EPUB_METADATA:: *setStr) (char* c,int length)=0;
	while (text[i] != 0) 
	{
		if (text[i] == '<')
		{
			if (state == STATE_NORMAL) 
			{
				int meta_pos = i;
				while (text[meta_pos] != ' '&& text[meta_pos] != '>')meta_pos++;
				if (CompareTag(_tagname_metadata, text, i + 1, meta_pos - 1 - i)) { state = STATE_IN_METADATA; i = meta_pos; continue; }//<metadata xmlns...>
			}
			if (state == STATE_IN_METADATA)
			{
				bool r = Element(text, i);
				if (r)
				{
					if (CompareTag(_tagname_creater, text, elementInfo.start + 1, elementInfo.tagname_length)) { setStr = &EPUB_METADATA::AddCreator; }
					else if (CompareTag(_tagname_title, text, elementInfo.start + 1, elementInfo.tagname_length)) { setStr = &EPUB_METADATA::SetTitle; }
					else if (CompareTag(_tagname_lang, text, elementInfo.start + 1, elementInfo.tagname_length)) { setStr = &EPUB_METADATA::SetLang; }
					if (setStr) { (meta.*setStr)(text + elementInfo.inner_start, elementInfo.inner_length); setStr = 0; 
					}
					i = elementInfo.end;
				}
				else 
				{
					if (text[i + 1] == '/')return S_OK;//</metadata>
				}
			}
		}
		i++;
	}

	
	return E_FAIL;
}



#pragma endregion