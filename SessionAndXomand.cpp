#include <Windows.h>
#include <oledb.h>
#include <oledberr.h>
#include <msdasc.h>
#include <tchar.h>
#include <strsafe.h>
#define COM_NO_WINDOWS_H
#define OLEDBVER 0x0260


#define DECLARE_BUFFER() TCHAR szBuf[1024] = {0};
#define OLEDB_PRINTF(...)\
	StringCchPrintf(szBuf, 1024, __VA_ARGS__);\
	WriteConsole(GetStdHandle(STD_OUTPUT_HANDLE), szBuf, _tcslen(szBuf), NULL, NULL);

#define DECLARE_OLEDB_INTERFACE(iid) iid *p##iid = NULL;

#define CHECK_OLEDB_INTERFACE(I, iid)\
	hRes = (I)->QueryInterface(IID_##iid, (void**)&(p##iid));\
	if(FAILED(hRes))\
	{\
		OLEDB_PRINTF(_T("不支持接口:%s\n"), _T(#iid));\
	}\
	else\
	{\
		OLEDB_PRINTF(_T("支持接口:%s\n"), _T(#iid));\
	}

#define OLEDB_SUCCESS(hr, ...)\
	if(FAILED(hr))\
	{\
		goto __CLEAR_UP;\
		OLEDB_PRINTF(__VA_ARGS__);\
	}
#define OLEDB_SAFE_RELEASE(p)\
	if((p) != NULL)\
	{\
		(p)->Release();\
	}


BOOL CreateDBSession(IOpenRowset* &pIOpenRowset)
{
	DECLARE_BUFFER();
	DECLARE_OLEDB_INTERFACE(IDBPromptInitialize);
	DECLARE_OLEDB_INTERFACE(IDBInitialize);
	DECLARE_OLEDB_INTERFACE(IDBCreateSession);
	BOOL bRet = FALSE;
	HWND hDesktop = GetDesktopWindow();
	HRESULT hRes = CoCreateInstance(CLSID_DataLinks, NULL, CLSCTX_INPROC_SERVER, IID_IDBPromptInitialize, (void**)&pIDBPromptInitialize);
	OLEDB_SUCCESS(hRes, _T("创建接口IDBPromptInitialize失败，错误码:%08x\n"), hRes);
	
	hRes = pIDBPromptInitialize->PromptDataSource(NULL, hDesktop, DBPROMPTOPTIONS_WIZARDSHEET, 0, NULL, NULL, IID_IDBInitialize, (IUnknown**)&pIDBInitialize);
	OLEDB_SUCCESS(hRes, _T("弹出接口数据源配置对话框失败，错误码:%08x\n"), hRes);
	
	hRes = pIDBInitialize->Initialize();
	OLEDB_SUCCESS(hRes, _T("链接数据库失败，错误码:%08x\n"), hRes);

	hRes = pIDBInitialize->QueryInterface(IID_IDBCreateSession, (void**)&pIDBCreateSession);

	OLEDB_SUCCESS(hRes, _T("创建接口IDBCreateSession失败，错误码:%08x\n"), hRes);
	hRes = pIDBCreateSession->CreateSession(NULL, IID_IOpenRowset, (IUnknown**)&pIOpenRowset);
	OLEDB_SUCCESS(hRes, _T("创建接口IOpenRowset失败，错误码:%08x\n"), hRes);
	bRet = TRUE;

__CLEAR_UP:
	OLEDB_SAFE_RELEASE(pIDBPromptInitialize);
	OLEDB_SAFE_RELEASE(pIDBInitialize);
	OLEDB_SAFE_RELEASE(pIDBCreateSession);
	return bRet;
}

int _tmain(int argc, TCHAR *argv[])
{
	DECLARE_BUFFER();
	DECLARE_OLEDB_INTERFACE(IOpenRowset);
	DECLARE_OLEDB_INTERFACE(IDBInitialize);
	DECLARE_OLEDB_INTERFACE(IDBCreateCommand);

	DECLARE_OLEDB_INTERFACE(IAccessor);
	DECLARE_OLEDB_INTERFACE(IColumnsInfo);
	DECLARE_OLEDB_INTERFACE(ICommand);
	DECLARE_OLEDB_INTERFACE(ICommandProperties);
	DECLARE_OLEDB_INTERFACE(ICommandText);
	DECLARE_OLEDB_INTERFACE(IConvertType);
	DECLARE_OLEDB_INTERFACE(IColumnsRowset);
	DECLARE_OLEDB_INTERFACE(ICommandPersist);
	DECLARE_OLEDB_INTERFACE(ICommandPrepare);
	DECLARE_OLEDB_INTERFACE(ICommandWithParameters);
	DECLARE_OLEDB_INTERFACE(ISupportErrorInfo);
	DECLARE_OLEDB_INTERFACE(ICommandStream);
	
	CoInitialize(NULL);
	if (!CreateDBSession(pIOpenRowset))
	{
		OLEDB_PRINTF(_T("调用函数CreateDBSession失败，程序即将推出\n"));
		return 0;
	}
	
	HRESULT hRes = pIOpenRowset->QueryInterface(IID_IDBCreateCommand, (void**)&pIDBCreateCommand);
	OLEDB_SUCCESS(hRes, _T("创建接口IDBCreateCommand失败，错误码:%08x\n"), hRes);
	
	hRes = pIDBCreateCommand->CreateCommand(NULL, IID_IAccessor, (IUnknown**)&pIAccessor);
	OLEDB_SUCCESS(hRes, _T("创建接口IAccessor失败，错误码:%08x\n"), hRes);

	CHECK_OLEDB_INTERFACE(pIAccessor, IColumnsInfo);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommand);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommandProperties);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommandText);
	CHECK_OLEDB_INTERFACE(pIAccessor, IConvertType);
	CHECK_OLEDB_INTERFACE(pIAccessor, IColumnsRowset);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommandPersist);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommandPrepare);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommandWithParameters);
	CHECK_OLEDB_INTERFACE(pIAccessor, ISupportErrorInfo);
	CHECK_OLEDB_INTERFACE(pIAccessor, ICommandStream);
	

__CLEAR_UP:
	OLEDB_SAFE_RELEASE(pIOpenRowset);
	OLEDB_SAFE_RELEASE(pIDBInitialize);
	OLEDB_SAFE_RELEASE(pIDBCreateCommand);

	OLEDB_SAFE_RELEASE(pIAccessor);
	OLEDB_SAFE_RELEASE(pIColumnsInfo);
	OLEDB_SAFE_RELEASE(pICommand);
	OLEDB_SAFE_RELEASE(pICommandProperties);
	OLEDB_SAFE_RELEASE(pICommandText);
	OLEDB_SAFE_RELEASE(pIConvertType);
	OLEDB_SAFE_RELEASE(pIColumnsRowset);
	OLEDB_SAFE_RELEASE(pICommandPersist);
	OLEDB_SAFE_RELEASE(pICommandPrepare);
	OLEDB_SAFE_RELEASE(pICommandWithParameters);
	OLEDB_SAFE_RELEASE(pISupportErrorInfo);
	OLEDB_SAFE_RELEASE(pICommandStream);
	CoUninitialize();
	return 0;
}
