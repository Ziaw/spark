
#pragma once

#include "atlutil.h"
#include "SparkLanguagePackage_i.h"

// EDCC3B79-0BAD-11DB-BC1A-00112FDE8B61
static const GUID NemerleLangGuid = 
		{ 0xEDCC3B79, 0x0BAD, 0x11DB, { 0xBC, 0x1A, 0x00, 0x11, 0x2F, 0xDE, 0x8B, 0x61 } };
static const GUID NemerleProjectTypeGuid = 
		{ 0xEDCC3B85, 0x0BAD, 0x11DB, { 0xBC, 0x1A, 0x00, 0x11, 0x2F, 0xDE, 0x8B, 0x61 } };


class SourceInit
{
public:
	CComPtr<IServiceProvider> _site;
	CComPtr<IVsTextLines> _primaryBuffer;
};

class ATL_NO_VTABLE Source :
	public CComCreatableObject<Source, SourceInit>,
	public ISparkSource,
	public IVsContainedLanguageHost,
	public ISourceSupervisorEvents
{
	CComPtr<ISourceSupervisor> _supervisor;
	DWORD _supervisorAdvise;

	CComPtr<IVsHierarchy> _hierarchy;
	VSITEMID _itemid;

	CComPtr<IVsTextLines> _secondaryBuffer;
	CComPtr<IVsTextBufferCoordinator> _bufferCoordinator;

	CComPtr<IVsIntellisenseProjectManager> _projectManager;
	CComPtr<IVsContainedLanguage> _containedLanguage;

	CComBSTR _primaryText;


	int _paintLength;
	SourcePainting* _paintArray;
	short _isNemerle;
	BOOL isNemerle() { BOOL test; IsNemerle(&test); return test; };

public:
	Source()
	{
		_supervisorAdvise = 0;
		_paintLength = 0;
		_paintArray = NULL;
		_isNemerle = -1;
	}


	BEGIN_COM_MAP(Source)
		COM_INTERFACE_ENTRY(ISparkSource)
		COM_INTERFACE_ENTRY(IVsContainedLanguageHost)
		COM_INTERFACE_ENTRY(ISourceSupervisorEvents)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT();

	HRESULT FinalConstruct();
	
	/**** ISparkSource ****/
	STDMETHODIMP GetSupervisor(ISourceSupervisor** ppSupervisor) {return _supervisor.CopyTo(ppSupervisor);}
	STDMETHODIMP SetSupervisor(ISourceSupervisor* pSupervisor);

	STDMETHODIMP GetIntellisenseProjectManager(IVsIntellisenseProjectManager** ppProjectManager)
	{
		if (_projectManager == NULL)
			return *ppProjectManager = NULL, S_OK;

		return _projectManager->QueryInterface(ppProjectManager);
	}

    STDMETHODIMP GetContainedLanguage( 
        /* [out] */ IVsContainedLanguage **ppContainedLanguage)
	{
		if (_containedLanguage == NULL)
			return *ppContainedLanguage = NULL, S_OK;

		return _containedLanguage->QueryInterface(ppContainedLanguage);
	}

    STDMETHODIMP GetTextBufferCoordinator( 
    /* [out] */ IVsTextBufferCoordinator **ppCoordinator)
	{
		if (_bufferCoordinator == NULL)
			return *ppCoordinator = NULL, S_OK;

		return _bufferCoordinator->QueryInterface(ppCoordinator);
	}

	STDMETHODIMP GetDocument(
		IVsHierarchy** ppHierarchy, 
		VSITEMID* pItemId, 
		IVsTextLines** pBuffer)
	{
		HRESULT hr = S_OK;
		if (ppHierarchy != NULL)
			_HR(_hierarchy.CopyTo(ppHierarchy));
		if (pItemId != NULL)
			*pItemId = _itemid;
		if (pBuffer != NULL)
			_HR(_primaryBuffer.CopyTo(pBuffer));
		return hr;
	}

    STDMETHODIMP GetPrimaryText(BSTR *pText)
	{
		*pText = _primaryText.Copy();
		return S_OK;
	}

	STDMETHODIMP GetRunningDocumentText(BSTR CanonicalName, BSTR *pText);

    STDMETHODIMP GetPaint( 
        /* [out] */ long *cPaint,
        /* [size_is][size_is][out] */ SourcePainting **prgPaint);

	STDMETHODIMP GetDefaultPageBaseType(BSTR* pPageBaseType);


	/**** ISourceSupervisorEvents ****/
    STDMETHODIMP OnGenerated( 
        /* [in] */ BSTR primaryText,
        /* [in] */ BSTR secondaryText,
        /* [in] */ long cMappings,
        /* [size_is][in] */ SourceMapping *rgSpans,
		/* [in] */ long cPaints,
		/* [size_is][in] */ SourcePainting *rgPaints);

	/**** IVsContainedLanguageHost ****/
	STDMETHODIMP Advise( 
		/* [in] */ __RPC__in_opt IVsContainedLanguageHostEvents *pHost,
		/* [out] */ __RPC__out VSCOOKIE *pvsCookie) {ATLTRACENOTIMPL(_T("Source::Advise"));}

	STDMETHODIMP Unadvise( 
		/* [in] */ VSCOOKIE vsCookie) {ATLTRACENOTIMPL(_T("Source::Unadvise"));}

	STDMETHODIMP GetLineIndent( 
		/* [in] */ long lLineNumber,
		/* [out] */ __RPC__deref_out_opt BSTR *pbstrIndentString,
		/* [out] */ __RPC__out long *plParentIndentLevel,
		/* [out] */ __RPC__out long *plIndentSize,
		/* [out] */ __RPC__out BOOL *pfTabs,
		/* [out] */ __RPC__out long *plTabSize);

	STDMETHODIMP CanReformatCode( 
		/* [out] */ __RPC__out BOOL *pfCanReformat) {ATLTRACENOTIMPL(_T("Source::CanReformatCode"));}

	STDMETHODIMP GetNearestVisibleToken( 
		/* [in] */ TextSpan tsSecondaryToken,
		/* [out] */ __RPC__out TextSpan *ptsPrimaryToken) 
	{
		ATLTRACE(_T("Source::GetNearestVisibleToken\r\n"));
		*ptsPrimaryToken = tsSecondaryToken;
		return S_OK;
		//ATLTRACENOTIMPL(_T("Source::GetNearestVisibleToken"));
	}

	STDMETHODIMP EnsureSpanVisible( 
		/* [in] */ TextSpan tsPrimary) {ATLTRACENOTIMPL(_T("Source::EnsureSpanVisible"));}

	STDMETHODIMP QueryEditFile() {ATLTRACENOTIMPL(_T("Source::QueryEditFile"));}

	STDMETHODIMP OnRenamed( 
		/* [in] */ ContainedLanguageRenameType clrt,
		/* [in] */ __RPC__in BSTR bstrOldID,
		/* [in] */ __RPC__in BSTR bstrNewID) {ATLTRACENOTIMPL(_T("Source::OnRenamed"));}

	STDMETHODIMP InsertControl( 
		/* [in] */ __RPC__in const WCHAR *pwcFullType,
		/* [in] */ __RPC__in const WCHAR *pwcID) {ATLTRACENOTIMPL(_T("Source::InsertControl"));}

	STDMETHODIMP InsertReference( 
		/* [in] */ __RPC__in const WCHAR *__MIDL__IVsContainedLanguageHost0000) {ATLTRACENOTIMPL(_T("Source::InsertReference"));}

	STDMETHODIMP GetVSHierarchy( 
		/* [out] */ __RPC__deref_out_opt IVsHierarchy **ppVsHierarchy) 
	{
		return _hierarchy->QueryInterface(ppVsHierarchy);
	}

	STDMETHODIMP GetErrorProviderInformation( 
		/* [out] */ __RPC__deref_out_opt BSTR *pbstrTaskProviderName,
		/* [out] */ __RPC__out GUID *pguidTaskProviderGuid) {ATLTRACENOTIMPL(_T("Source::GetErrorProviderInformation"));}

	STDMETHODIMP InsertImportsDirective( 
		/* [in] */ __RPC__in const WCHAR *__MIDL__IVsContainedLanguageHost0001) {ATLTRACENOTIMPL(_T("Source::InsertImportsDirective"));}

	STDMETHODIMP OnContainedLanguageEditorSettingsChange() {ATLTRACENOTIMPL(_T("Source::OnContainedLanguageEditorSettingsChange"));}

	STDMETHODIMP EnsureSecondaryBufferReady();

	STDMETHODIMP IsNemerle(__RPC__out BOOL *pFlag);
};

