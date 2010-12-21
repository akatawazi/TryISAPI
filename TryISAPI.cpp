// TRYISAPI.CPP - Implementation file for your Internet Server
//    TryISAPI Filter

#include "stdafx.h"
#include "TryISAPI.h"
#include <iostream> 
#include <fstream> 
#import "msado15.dll" \
	no_namespace \
	rename( "EOF", "adoEOF" )
#define DEFAULT_BUFFER_SIZE         4096
#define WRITE_BUFFER_SIZE   1000 
#define MODULE_NAME         "HTTPLogger" 

#if !defined(uniqueID)
#define uniqueID
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER


#endif // uniqueID

///////////////////////////////////////////////////////////////////////
// The one and only CWinApp object
// NOTE: You may remove this object if you alter your project to no
// longer use MFC in a DLL.

CWinApp theApp;
CString mDBConnectString;
CString mServerName;
const char* szDBConnectString = "DB Connection String";
const char* szServerName = "Server Name";
const char* szTraceLevel = "Trace Level";
const char* szOutputLevel = "Output Level";
const char* szDBConnectStringEnd = "<DB Connection String>";
const char* szServerNameEnd = "<Server Name>";
const char* szTraceLevelEnd = "<Trace Level>";
const char* szOutputLevelEnd = "<Output Level>";
const CString INIfileName = "d:\\WebSites\\HTTPLogger\\ISAPI.ini";
const char* outFileName = "D:\\WebSites\\HTTPLogger\\test.txt";
int mTraceLevel;
int mLoggingLevel; 
BOOL ReadConfigurationFile();
BOOL ReadServerName(CStdioFile& file);
BOOL ReadTraceLevel(CStdioFile& file);
BOOL ReadOutputLevel(CStdioFile& file);
BOOL ReadDBConfiguration(CStdioFile& file);
//std::string RequestInsert(std::string URL, std::string IP, CString ServerName, std::string Cookie, std::string Lang, std::string UAgent);
CString RequestInsert(std::string URL, std::string RawHeader, CString ServerName, std::string QueryString);
CString CrackStrVariant(const _variant_t& var);
VOID WriteDebug(LPSTR   szFormat, ... );
VOID ResponseInsert(LPSTR transactionID, LPSTR status);
///////////////////////////////////////////////////////////////////////
// The one and only CTryISAPIFilter object

CTryISAPIFilter theFilter;


///////////////////////////////////////////////////////////////////////
// CTryISAPIFilter implementation

CTryISAPIFilter::CTryISAPIFilter()
{
}

CTryISAPIFilter::~CTryISAPIFilter()
{
}

BOOL CTryISAPIFilter::GetFilterVersion(PHTTP_FILTER_VERSION pVer)
{
	// Call default implementation for initialization
	CHttpFilter::GetFilterVersion(pVer);

	// Clear the flags set by base class
	pVer->dwFlags &= ~SF_NOTIFY_ORDER_MASK;

	pVer->dwFlags |= SF_NOTIFY_ORDER_LOW | SF_NOTIFY_SECURE_PORT | SF_NOTIFY_NONSECURE_PORT	 | SF_NOTIFY_PREPROC_HEADERS | SF_NOTIFY_SEND_RAW_DATA | SF_NOTIFY_SEND_RESPONSE;

	// Load description string
	TCHAR sz[SF_MAX_FILTER_DESC_LEN+1];
	ISAPIVERIFY(::LoadString(AfxGetResourceHandle(),
			IDS_FILTER, sz, SF_MAX_FILTER_DESC_LEN));
	_tcscpy(pVer->lpszFilterDesc, sz);
	
	return ReadConfigurationFile();
}

DWORD CTryISAPIFilter::OnPreprocHeaders(CHttpFilterContext* pCtxt,	PHTTP_FILTER_PREPROC_HEADERS pHeaderInfo)
{
	if(mLoggingLevel==1){
		BOOL                            fRet = FALSE;
		CHAR                            pBuf[ DEFAULT_BUFFER_SIZE ];
		CHAR *                          pszBuf = pBuf;
		CHAR                            pBuf2[ DEFAULT_BUFFER_SIZE ];
		CHAR *                          pszBuf2 = pBuf2;
		DWORD                           cbBuf2 = DEFAULT_BUFFER_SIZE;
		CHAR                            pBuf3[ DEFAULT_BUFFER_SIZE ];
		CHAR *                          pszBuf3 = pBuf3;
		DWORD                           cbBuf3 = DEFAULT_BUFFER_SIZE;
		
		WriteDebug("OnPreprocHeaders()");
		fRet = pCtxt->GetServerVariable("ALL_RAW", pszBuf2, &cbBuf2 );
		WriteDebug("ALL_RAW: %s",pszBuf2);
		fRet = pCtxt->GetServerVariable("URL", pszBuf, &cbBuf2 );
		WriteDebug("URL: %s\r\n",pszBuf);
		fRet = pCtxt->GetServerVariable("QUERY_STRING", pszBuf3, &cbBuf3 );
		WriteDebug("QUERY_STRING: %s\r\n",pszBuf3);
		CString PKID = "PKID: ";
		CString ret_pkid = RequestInsert(pszBuf, pszBuf2, mServerName, pszBuf3);
		WriteDebug("Got back from RequestInsert: %s\r\n",ret_pkid);
		PKID.Append(ret_pkid);
		PKID.Append("\r\n");
		WriteDebug("PKID = %s\r\n",PKID);
		LPSTR LPPKID = PKID.GetBuffer(PKID.GetLength());
		WriteDebug("Setting Response Header: %s\r\n",LPPKID);
		pCtxt->AddResponseHeaders(LPPKID,0);
	}
	return SF_STATUS_REQ_NEXT_NOTIFICATION;

}
DWORD CTryISAPIFilter::OnSendResponse(CHttpFilterContext * pfc, PHTTP_FILTER_SEND_RESPONSE  pSR)
{
	if(mLoggingLevel==1){
		char dateStr [9];
		char timeStr [9];
		_strdate( dateStr);
		_strtime( timeStr );

		BOOL                            fRet = FALSE;
		CHAR                            pBuf[ DEFAULT_BUFFER_SIZE ];
		CHAR *                          pszBuf = pBuf;
		CHAR                            pBuf2[ DEFAULT_BUFFER_SIZE ];
		CHAR *                          pszBuf2 = pBuf2;
		DWORD                           cbBuf = DEFAULT_BUFFER_SIZE;
		DWORD                           cbBuf2 = DEFAULT_BUFFER_SIZE;
		CHAR                            szPKID[] = "PKID:";
		CHAR                            szPKID2[] = "status";

		fRet = pSR->GetHeader( pfc->m_pFC, szPKID, pszBuf, &cbBuf );

		fRet = pSR->GetHeader( pfc->m_pFC, szPKID2, pszBuf2, &cbBuf2 );

		ResponseInsert(pszBuf, pszBuf2);
	}
	return SF_STATUS_REQ_NEXT_NOTIFICATION;
}
VOID ResponseInsert(LPSTR transactionID, LPSTR status)
{
	WriteDebug("Begin ResponseInsert");
	WriteDebug("TransactionID: %s\r\n",transactionID);
	WriteDebug("status: %s\r\n",status);
	try{
		_ConnectionPtr m_pConn;
		//create the COM database object:
		HRESULT hr = m_pConn.CreateInstance(__uuidof (Connection));

		if(FAILED(hr))
		{
			WriteDebug("instantiate COM object failed\r\n");
		}
		WriteDebug("success instantiate COM object\r\n");
		WriteDebug("attempt connect to TMS DB\r\n");
		//connect to the database:

		hr = m_pConn->Open(_bstr_t(mDBConnectString), _bstr_t(""), _bstr_t(""), adModeUnknown);
		if(FAILED(hr))
		{
			WriteDebug("connect to TMS DB failed\r\n");
		}
		WriteDebug("success connect to TMS DB\r\n");

		_CommandPtr pCommand;
		_ParameterPtr pParam;
		WriteDebug("attempt create SQL command object\r\n");
		//create a database command object:
		hr = pCommand.CreateInstance(__uuidof(Command));

		if(FAILED(hr))
		{
			WriteDebug("create SQL command object failed\r\n");
		}
		WriteDebug("success create SQL command object\r\n");
		CString m_strSCN;
		pCommand->ActiveConnection = m_pConn;
		pCommand->CommandText = "HTTPLogger.HTTPLogger_owner.ResponseInsert";
		pCommand->CommandType = adCmdStoredProc;

		//set the arguments:
		WriteDebug("attempt set sproc parameters\r\n");
		pParam = pCommand->CreateParameter("@Status",adVarChar, adParamInput, 1000, (LPCTSTR)status);      
		pCommand->Parameters->Append(pParam);
		pParam = pCommand->CreateParameter("@TransactionID",adVarChar, adParamInput, 200, (LPCTSTR)transactionID);      
		pCommand->Parameters->Append(pParam);

		WriteDebug("success set sproc parameters\r\n");

		WriteDebug("attempt execute sproc\r\n");
		//execute the stored procedure:
		pCommand->Execute( NULL,NULL,adCmdStoredProc );
		
		WriteDebug("success execute sproc\r\n");
		
		WriteDebug("attempt close connection");
		m_pConn->Close();
		WriteDebug("success close connection");
	}
	catch(_com_error &e)
	{
		_bstr_t bstrSource(e.Source());
		_bstr_t bstrDescription(e.Description());
		WriteDebug("caught COM error\r\n");
		WriteDebug(e.Source());
		WriteDebug(e.Description());
	}
	catch(...)
	{
		WriteDebug("caught some other error\r\n");
	}
	WriteDebug("end of ResponseInsert()\r\n");
}

BOOL ReadConfigurationFile(){
	/*
	* This function reads the configuration values form the .ini file.
	* Note the hard-coded path and file name.
	*/

	CStdioFile file;
	CFileException ex;
	//open the file:
	if(!file.Open(INIfileName, CFile::modeRead, &ex))
	{
		char szError[256];
		ex.GetErrorMessage(szError, 256);
		//WriteDebug(0,"Error reading .ini file: %s\n",szError);
		return FALSE;
	}

	try
	{
		CString str;
		//once the opening bracket for a section is found, continue to read until the
		//ending bracket is found.  Improperly formatted .ini files will cause
		//this function to fail, or perhaps run forever.
		while(file.ReadString(str))
		{
			if(str == szDBConnectString && !ReadDBConfiguration(file))
				break;
			else if(str == szServerName && !ReadServerName(file))
				break;
			else if(str == szTraceLevel && !ReadTraceLevel(file))
				break;
			else if(str == szOutputLevel && !ReadOutputLevel(file))
				break;
		}
		if(mDBConnectString.GetLength() == 0)
		{
			//WriteDebug(0,"Configuration file is not in expected format\n");
			return FALSE;
		}
	}
	catch(CFileException* e)
	{
		char szError[256];
		e->GetErrorMessage(szError, 256);
		//WriteDebug(0,"Exception caught: %s\n",szError);
		return FALSE;
	}
	file.Close();
	return TRUE;
}
BOOL ReadDBConfiguration(CStdioFile& file)
{
	/* This function reads from the passed-in file handle, until the ending bracket
	* for that section is encountered.
	*/
	CString str;
	BOOL bReturn = FALSE;

	while(file.ReadString(str))
	{
		if(str == szDBConnectStringEnd)
		{
			bReturn = TRUE;
			break;
		}
		mDBConnectString = str;
	}
	//WriteDebug(0,"Got DB Connection String: %s\n",mDBConnectString);
	return bReturn;
}
BOOL ReadServerName(CStdioFile& file)
{
	/* This function reads from the passed-in file handle, until the ending bracket
	* for that section is encountered.
	*/
	CString str;
	BOOL bReturn = FALSE;

	while(file.ReadString(str))
	{
		if(str == szServerNameEnd)
		{
			bReturn = TRUE;
			break;
		}
		mServerName = str;
	}
	//WriteDebug(0,"Got DB Connection String: %s\n",mDBConnectString);
	return bReturn;
}
BOOL ReadTraceLevel(CStdioFile& file)
{
	/* This function reads from the passed-in file handle, until the ending bracket
	* for that section is encountered.
	*/
	CString str;
	BOOL bReturn = FALSE;

	while(file.ReadString(str))
	{
		if(str == szTraceLevelEnd)
		{
			bReturn = TRUE;
			break;
		}
		mTraceLevel = atoi(str);
	}
	//WriteDebug(0,"Got DB Connection String: %s\n",mDBConnectString);
	return bReturn;
}

BOOL ReadOutputLevel(CStdioFile& file)
{
	/* This function reads from the passed-in file handle, until the ending bracket
	* for that section is encountered.
	*/
	CString str;
	BOOL bReturn = FALSE;

	while(file.ReadString(str))
	{
		if(str == szOutputLevelEnd)
		{
			bReturn = TRUE;
			break;
		}
		mLoggingLevel = atoi(str);
	}
	//WriteDebug(0,"Got DB Connection String: %s\n",mDBConnectString);
	return bReturn;
}
CString RequestInsert(std::string URL, std::string RawHeader, CString ServerName, std::string QueryString)
{
	BOOL Return = FALSE;

	//Establishing a connection to the datasource
	try	
	{
		//replace all carriage returns with '/n'
		char cr = 10;
		int position = RawHeader.find( cr ); // find first space
		
		WriteDebug("begin replace line feeds in RawHeader");
		// 
		while ( position != std::string::npos ) 
		{
			RawHeader.replace( position, 1, "|||" );
			position = RawHeader.find( cr, position + 1 );
		} 
		WriteDebug("end replace line feeds in RawHeader");

		_ConnectionPtr m_pConn;
		//create the COM database object:
		HRESULT hr = m_pConn.CreateInstance(__uuidof (Connection));

		if(FAILED(hr))
		{
			WriteDebug("instantiate COM object failed\r\n");
			return "";
		}
		WriteDebug("success instantiate COM object\r\n");
		WriteDebug("attempt connect to TMS DB\r\n");
		//connect to the database:

		hr = m_pConn->Open(_bstr_t(mDBConnectString), _bstr_t(""), _bstr_t(""), adModeUnknown);
		if(FAILED(hr))
		{
			WriteDebug("connect to TMS DB failed\r\n");
			return "";
		}
		WriteDebug("success connect to TMS DB\r\n");

		_CommandPtr pCommand;
		_ParameterPtr pParam;
		WriteDebug("attempt create SQL command object\r\n");
		//create a database command object:
		hr = pCommand.CreateInstance(__uuidof(Command));

		if(FAILED(hr))
		{
			WriteDebug("create SQL command object failed\r\n");
			return "";
		}
		WriteDebug("success create SQL command object\r\n");
		CString m_strSCN;
		pCommand->ActiveConnection = m_pConn;
		pCommand->CommandText = "HTTPLogger.HTTPLogger_owner.RequestInsert";
		pCommand->CommandType = adCmdStoredProc;

		//set the arguments:
		WriteDebug("attempt set sproc parameters\r\n");
		CString mURL = URL.c_str();
		pParam = pCommand->CreateParameter("@URL",adVarChar, adParamInput, 4000, (LPCTSTR)mURL);      
		pCommand->Parameters->Append(pParam);
		CString mRawHeader = RawHeader.c_str();
		pParam = pCommand->CreateParameter("@RawHeader",adVarChar, adParamInput, 4000, (LPCTSTR)mRawHeader);      
		pCommand->Parameters->Append(pParam);
		pParam = pCommand->CreateParameter("@Servername",adVarChar, adParamInput, 50, (LPCTSTR)ServerName);      
		pCommand->Parameters->Append(pParam);
		CString mQueryString = QueryString.c_str();
		pParam = pCommand->CreateParameter("@QueryString",adVarChar, adParamInput, 4000, (LPCTSTR)mQueryString);      
		pCommand->Parameters->Append(pParam);

		WriteDebug("success set sproc parameters\r\n");

		WriteDebug("attempt execute sproc\r\n");
		//execute the stored procedure:
		_RecordsetPtr Rs1 = pCommand->Execute( NULL,NULL,adCmdStoredProc );
		
		WriteDebug("success execute sproc\r\n");
		WriteDebug("attempt retrieve output parameter1\r\n");
		
		//get the output value and return it:
		_variant_t Final;		
		CString strTmp;			
		WriteDebug("attempt retrieve output parameter2\r\n");
		Final  = Rs1->Fields->GetItem( _variant_t( 0L ) )->Value;		
		WriteDebug("attempt retrieve output parameter3\r\n");
		strTmp.Format( "%s", CrackStrVariant( Final) );
		WriteDebug("success retrieve output parameter: %s\r\n",strTmp);
		WriteDebug("attempt close connection");
		m_pConn->Close();
		WriteDebug("success close connection");

		return strTmp;
	}
	catch(_com_error &e)
	{
		_bstr_t bstrSource(e.Source());
		_bstr_t bstrDescription(e.Description());
		WriteDebug("caught COM error\r\n");
		WriteDebug(e.Source());
		WriteDebug(e.Description());
	}
	catch(...)
	{
		WriteDebug("caught some other error\r\n");
	}
	WriteDebug("end of RequestInsert()\r\n");
	return "";
}
CString CrackStrVariant(const _variant_t& var)
{
	/* This code was pulled whole off of the web.
	* It takes in a variant, determines the data type,, and returns it as a CString.
	*/
    CString strRet;
    strRet = _T("<Unknown>");
    switch(var.vt)
     {
          case VT_EMPTY:
          case VT_NULL:
              strRet = _T("NULL");
			  //WriteDebug(SID,"type null\n");
               break;
          case VT_I2:
               strRet.Format(_T("%hd"), V_I2(&var));
			   //WriteDebug(SID,"type i2\n");
               break;
          case VT_I4:
               strRet.Format(_T("%d"),V_I4(&var));
			   //WriteDebug(SID,"type i4\n");
               break;
          case VT_R4:
               strRet.Format(_T("%e"), (double)V_R4(&var));
			   //WriteDebug(SID,"type r4\n");
               break;
          case VT_R8:
               strRet.Format(_T("%e"), V_R8(&var));
			   //WriteDebug(SID,"type r8\n");
               break;
          case VT_BSTR:
               strRet = V_BSTR(&var);
			   //WriteDebug(SID,"type bstr\n");
               break;
          case VT_DISPATCH:
               strRet = _T("VT_DISPATCH");
			   //WriteDebug(SID,"type dispatch\n");
               break;
          case VT_ERROR:
               strRet = _T("VT_ERROR");
			   //WriteDebug(SID,"type error\n");
               break;
          case VT_BOOL:
			  //WriteDebug(SID,"type bool\n");
               return ( V_BOOL(&var) ? _T("TRUE") : _T("FALSE"));
          case VT_VARIANT:
               strRet = _T("VT_VARIANT");
			   //WriteDebug(SID,"type variant\n");
               break;
          case VT_UNKNOWN:
               strRet = _T("VT_UNKNOWN");
			   //WriteDebug(SID,"type unknown\n");
               break;
          case VT_I1:
               strRet = _T("VT_I1");
			   //WriteDebug(SID,"type i1\n");
               break;
          case VT_UI1:
               strRet.Format(_T("0x%02hX"), (unsigned short)V_UI1(&var));
			   //WriteDebug(SID,"type ui1\n");
               break;
          case VT_UI2:
               strRet = _T("VT_UI2");
			   //WriteDebug(SID,"type ui2\n");
               break;
          case VT_UI4:
               strRet = _T("VT_UI4");
			   //WriteDebug(SID,"type ui4\n");
               break;
          case VT_I8:
               strRet = _T("VT_I8");
			   //WriteDebug(SID,"type i8\n");
               break;
          case VT_UI8:
               strRet = _T("VT_UI8");
			   //WriteDebug(SID,"type ui8\n");
               break;
          case VT_INT:
               strRet = _T("VT_INT");
			   //WriteDebug(SID,"type int\n");
               break;
          case VT_UINT:
               strRet = _T("VT_UINT");
			   //WriteDebug(SID,"type uint\n");
               break;
          case VT_VOID:
               strRet = _T("VT_VOID");
			   //WriteDebug(SID,"type void\n");
               break;
          case VT_HRESULT:
               strRet = _T("VT_HRESULT");
			   //WriteDebug(SID,"type hresult\n");
               break;
          case VT_PTR:
               strRet = _T("VT_PTR");
			   //WriteDebug(SID,"type ptr\n");
               break;
          case VT_SAFEARRAY:
               strRet = _T("VT_SAFEARRAY");
			   //WriteDebug(SID,"type safearray\n");
               break;
          case VT_CARRAY:
               strRet = _T("VT_CARRAY");
			   //WriteDebug(SID,"type carray\n");
               break;
          case VT_USERDEFINED:
               strRet = _T("VT_USERDEFINED");
			   //WriteDebug(SID,"type userdefined\n");
               break;
          case VT_LPSTR:
               strRet = _T("VT_LPSTR");
			   //WriteDebug(SID,"type lpstr\n");
               break;
          case VT_LPWSTR:
               strRet = _T("VT_LPWSTR");
			   //WriteDebug(SID,"type lpwstr\n");
               break;
          case VT_FILETIME:
               strRet = _T("VT_FILETIME");
			   //WriteDebug(SID,"type filetime\n");
               break;
          case VT_BLOB:
               strRet = _T("VT_BLOB");
			   //WriteDebug(SID,"type blob\n");
               break;
          case VT_STREAM:
               strRet = _T("VT_STREAM");
			   //WriteDebug(SID,"type stream\n");
               break;
          case VT_STORAGE:
               strRet = _T("VT_STORAGE");
			   //WriteDebug(SID,"type storage\n");
               break;
          case VT_STREAMED_OBJECT:
               strRet = _T("VT_STREAMED_OBJECT");
			   //WriteDebug(SID,"type streamedobject\n");
               break;
          case VT_STORED_OBJECT:
               strRet = _T("VT_STORED_OBJECT");
			   //WriteDebug(SID,"type storedobject\n");
               break;
          case VT_BLOB_OBJECT:
               strRet = _T("VT_BLOB_OBJECT");
			   //WriteDebug(SID,"type blobobject\n");
               break;
          case VT_CF:
               strRet = _T("VT_CF");
			   //WriteDebug(SID,"type cf\n");
               break;
          case VT_CLSID:
               strRet = _T("VT_CLSID");
			   //WriteDebug(SID,"type clsid\n");
               break;
    }
   
     WORD vt = var.vt;
    if(vt & VT_ARRAY)
     {
        vt = vt & ~VT_ARRAY;
        strRet = _T("Array of ");
    }
   
     if(vt & VT_BYREF)
     {
        vt = vt & ~VT_BYREF;
        strRet += _T("Pointer to ");
    }
   
     if(vt != var.vt)
     {
        switch(vt)
          {
               case VT_EMPTY:
                    strRet += _T("VT_EMPTY");
                    break;
               case VT_NULL:
                    strRet += _T("VT_NULL");
                    break;
               case VT_I2:
                    strRet += _T("VT_I2");
                    break;
               case VT_I4:
                    strRet += _T("VT_I4");
                    break;
               case VT_R4:
                    strRet += _T("VT_R4");
                    break;
               case VT_R8:
                    strRet += _T("VT_R8");
                    break;
               case VT_CY:
                    strRet += _T("VT_CY");
                    break;
               case VT_DATE:
                    strRet += _T("VT_DATE");
                    break;
               case VT_BSTR:
                    strRet += _T("VT_BSTR");
                    break;
               case VT_DISPATCH:
                    strRet += _T("VT_DISPATCH");
                    break;
               case VT_ERROR:
                    strRet += _T("VT_ERROR");
                    break;
               case VT_BOOL:
                    strRet += _T("VT_BOOL");
                    break;
               case VT_VARIANT:
                    strRet += _T("VT_VARIANT");
                    break;
               case VT_UNKNOWN:
                    strRet += _T("VT_UNKNOWN");
                    break;
               case VT_I1:
                    strRet += _T("VT_I1");
                    break;
               case VT_UI1:
                    strRet += _T("VT_UI1");
                    break;
               case VT_UI2:
                    strRet += _T("VT_UI2");
                    break;
               case VT_UI4:
                    strRet += _T("VT_UI4");
                    break;
               case VT_I8:
                    strRet += _T("VT_I8");
                    break;
               case VT_UI8:
                    strRet += _T("VT_UI8");
                    break;
               case VT_INT:
                    strRet += _T("VT_INT");
                    break;
               case VT_UINT:
                    strRet += _T("VT_UINT");
                    break;
               case VT_VOID:
                    strRet += _T("VT_VOID");
                    break;
               case VT_HRESULT:
                    strRet += _T("VT_HRESULT");
                    break;
               case VT_PTR:
                    strRet += _T("VT_PTR");
                    break;
               case VT_SAFEARRAY:
                    strRet += _T("VT_SAFEARRAY");
                    break;
               case VT_CARRAY:
                    strRet += _T("VT_CARRAY");
                    break;
               case VT_USERDEFINED:
                    strRet += _T("VT_USERDEFINED");
                    break;
               case VT_LPSTR:
                    strRet += _T("VT_LPSTR");
                    break;
               case VT_LPWSTR:
                    strRet += _T("VT_LPWSTR");
                    break;
               case VT_FILETIME:
                    strRet += _T("VT_FILETIME");
                    break;
               case VT_BLOB:
                    strRet += _T("VT_BLOB");
                    break;
               case VT_STREAM:
                    strRet += _T("VT_STREAM");
                    break;
               case VT_STORAGE:
                    strRet += _T("VT_STORAGE");
                    break;
               case VT_STREAMED_OBJECT:
                    strRet += _T("VT_STREAMED_OBJECT");
                    break;
               case VT_STORED_OBJECT:
                    strRet += _T("VT_STORED_OBJECT");
                    break;
               case VT_BLOB_OBJECT:
                    strRet += _T("VT_BLOB_OBJECT");
                    break;
               case VT_CF:
                    strRet += _T("VT_CF");
                    break;
               case VT_CLSID:
                    strRet += _T("VT_CLSID");
                    break;
        }
    }
   
     return strRet;
}
VOID WriteDebug(LPSTR   szFormat, ... ) { 
	/* 
	* This writes to a debug log file.
	* Please note that the log file path and name are hard-coded to
	* D:\WebSites\wwwroot\ISAPI.log
	*/
	if(mTraceLevel==1){
		CHAR    szBuffer[WRITE_BUFFER_SIZE]; 
		LPSTR   pCursor; 
		DWORD   cbToWrite; 
		INT     nWritten; 
		va_list args; 

		if ( WRITE_BUFFER_SIZE < 3 ) 
		{ 
			// 
			// This is just too small to deal with... 
			// 

			return; 
		} 

		// 
		// Inject the module name tag into the buffer 
		// 

		nWritten = _snprintf( szBuffer, 
							WRITE_BUFFER_SIZE, 
							"[%s.dll] ", 
							MODULE_NAME ); 

		if ( nWritten == -1 ) 
		{ 
			return; 
		} 

		pCursor = szBuffer + nWritten; 
		cbToWrite = WRITE_BUFFER_SIZE - nWritten; 

		va_start( args, szFormat ); 

		nWritten = _vsnprintf( pCursor, 
							cbToWrite, 
							szFormat, 
							args ); 

		va_end( args ); 

		if ( nWritten == -1 ) 
		{ 
			szBuffer[WRITE_BUFFER_SIZE-3] = '\r'; 
			szBuffer[WRITE_BUFFER_SIZE-2] = '\n'; 
		} 

		szBuffer[WRITE_BUFFER_SIZE-1] = '\0'; 

		char dateStr [9];
        char timeStr [9];
        _strdate( dateStr);
        _strtime( timeStr );
        
		std::ofstream os_out(outFileName,std::ios::app);
		os_out<<dateStr<<" "<<timeStr<<": "<<szBuffer<<"\n";
		os_out.close();
	}
} 
#if 0
BEGIN_MESSAGE_MAP(CTryISAPIFilter, CHttpFilter)
	//{{AFX_MSG_MAP(CTryISAPIFilter)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()
#endif	// 0
