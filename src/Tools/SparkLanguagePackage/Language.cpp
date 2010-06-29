
#include "stdafx.h"
#include "Language.h"
#include "CodeWindowManager.h"
#include "Colorizer.h"
#include "Source.h"
#include "ColorableItem.h"


STDMETHODIMP Language::GetSource(IVsTextBuffer* pBuffer, ISparkSource** ppSource)
{
	HRESULT hr = S_OK;

	CComCritSecLock<CComCriticalSection> lock(_sourcesLock);

	CComPtr<IUnknown> key;
	_HR(pBuffer->QueryInterface(&key));

	int index = _sources.FindKey(key);

	if (index == -1)
	{
		CComPtr<ISparkSource> source;
		
		SourceInit init = {_site};
		_HR(pBuffer->QueryInterface(&init._primaryBuffer));
		_HR(Source::CreateInstance(init, &source));
		
		_HR(_supervisor->OnSourceAssociated(source));

		_HR(source->QueryInterface(ppSource));

		if (SUCCEEDED(hr))
			_sources.Add(key.Detach(), source.Detach());
	}
	else
	{
		_sources.GetValueAt(index)->QueryInterface(ppSource);
	}
	return hr;
}

STDMETHODIMP Language::GetColorizer( 
    /* [in] */ __RPC__in_opt IVsTextLines *pBuffer,
    /* [out] */ __RPC__deref_out_opt IVsColorizer **ppColorizer)
{
	HRESULT hr = S_OK;

	if (_languageItems)
	{
		_languageItems.Detach();
		_languageItems = NULL;
	}

	// detect language
	// Get the moniker for the text buffer
	CComPtr<IVsUserData> userData;
	_HR(pBuffer->QueryInterface(&userData));		
	CComVariant moniker;
	_HR(userData->GetData(__uuidof(IVsUserData), &moniker));
	_HR(moniker.ChangeType(VT_BSTR));

	// Locate hierarchy itemid
	CComPtr<IWebApplicationCtxSvc> webApplicationCtx;
	_HR(_site->QueryService(__uuidof(IWebApplicationCtxSvc), &webApplicationCtx));

	CComPtr<IVsHierarchy> hierarchy;
	VSITEMID itemid;
	_HR(webApplicationCtx->GetItemContextFromPath(V_BSTR(&moniker), FALSE, &hierarchy, &itemid));

	GUID projectTypeGuid;
	_HR(hierarchy->GetGuidProperty(VSITEMID_ROOT, VSHPROPID_TypeGuid, &projectTypeGuid));

	BOOL isNemerle = projectTypeGuid == NemerleProjectTypeGuid;

	if (isNemerle)
	{
		_HR(_site->QueryService(NemerleLangGuid, &_languageItems))
	}
	else
	{
		_HR(_site->QueryService(__uuidof(CSharp), &_languageItems));
	}

	int languageItemCount;
	_HR(_languageItems->GetItemCount(&languageItemCount));

	ColorizerInit init = {this, pBuffer, languageItemCount};
	_HR(Colorizer::CreateInstance(init, ppColorizer));
	return hr;
}

STDMETHODIMP Language::GetCodeWindowManager( 
    /* [in] */ __RPC__in_opt IVsCodeWindow *pCodeWin,
    /* [out] */ __RPC__deref_out_opt IVsCodeWindowManager **ppCodeWinMgr)
{
	HRESULT hr = S_OK;
	CodeWindowManagerInit init = {this, pCodeWin};
	_HR(CodeWindowManager::CreateInstance(init, ppCodeWinMgr));
	return hr;
}


STDMETHODIMP Language::GetItemCount( 
    /* [out] */ __RPC__out int *piCount)
{
	HRESULT hr = S_OK;

	int languageItemCount;
	_HR(_languageItems->GetItemCount(&languageItemCount));

	CComPtr<IVsProvideColorableItems> sparkItems;
	_HR(_supervisor->QueryInterface(&sparkItems));
	int sparkItemCount;
	_HR(sparkItems->GetItemCount(&sparkItemCount));
	
	*piCount = languageItemCount + sparkItemCount;
	return hr;
}

STDMETHODIMP Language::GetColorableItem( 
    /* [in] */ int iIndex,
    /* [out] */ __RPC__deref_out_opt IVsColorableItem **ppItem)
{
	HRESULT hr = S_OK;

	int languageItemCount;
	_HR(_languageItems->GetItemCount(&languageItemCount));

	if (SUCCEEDED(hr) && iIndex <= languageItemCount)
	{
		_HR(_languageItems->GetColorableItem(iIndex, ppItem));
		
		CComBSTR name;
		_HR((*ppItem)->GetDisplayName(&name));

		return hr;
	}

	// return spark color info for upper band
	CComPtr<IVsProvideColorableItems> sparkItems;
	_HR(_supervisor->QueryInterface(&sparkItems));

	int sparkItemCount;
	_HR(sparkItems->GetItemCount(&sparkItemCount));

	if (SUCCEEDED(hr) && iIndex <= languageItemCount + sparkItemCount)
	{
		_HR(sparkItems->GetColorableItem(iIndex - languageItemCount, ppItem));

		CComBSTR name;
		_HR((*ppItem)->GetDisplayName(&name));

		return hr;
	}
	return E_INVALIDARG;
}
