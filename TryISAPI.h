#if !defined(AFX_TRYISAPI_H__B01D78A1_8BC4_4D02_BC24_22C79546A35B__INCLUDED_)
#define AFX_TRYISAPI_H__B01D78A1_8BC4_4D02_BC24_22C79546A35B__INCLUDED_

// TRYISAPI.H - Header file for your Internet Server
//    TryISAPI Filter

#include "resource.h"


class CTryISAPIFilter : public CHttpFilter
{
public:
	CTryISAPIFilter();
	~CTryISAPIFilter();

// Overrides
	// ClassWizard generated virtual function overrides
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//{{AFX_VIRTUAL(CTryISAPIFilter)
	public:
	virtual BOOL GetFilterVersion(PHTTP_FILTER_VERSION pVer);
	virtual DWORD OnPreprocHeaders(CHttpFilterContext* pCtxt, PHTTP_FILTER_PREPROC_HEADERS pHeaderInfo);
	//virtual DWORD OnSendResponse(HTTP_FILTER_CONTEXT *        pfc, HTTP_FILTER_SEND_RESPONSE *  pSR);
	virtual DWORD OnSendResponse(CHttpFilterContext *, PHTTP_FILTER_SEND_RESPONSE);
	//virtual DWORD OnSendRawData(CHttpFilterContext* pCtxt, PHTTP_FILTER_RAW_DATA pRawData);
	//virtual DWORD OnReadRawData(CHttpFilterContext* pCtxt, PHTTP_FILTER_RAW_DATA pRawData);
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CTryISAPIFilter)
	//}}AFX_MSG
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TRYISAPI_H__B01D78A1_8BC4_4D02_BC24_22C79546A35B__INCLUDED)
