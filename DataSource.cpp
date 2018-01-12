#include <windows.h>
#include <oledb.h>
#include <oledberr.h>
#include <tchar.h>
#include <strsafe.h>
#include <msdasc.h>

#define COM_NO_WINDOWS_H
#define OLEDBVER 0X0260 //定义OLEDB库版本
#define DECLARE_BUFFER() TCHAR szBuf[1024] = _T("")
#define COM_PRINTF(...)\
	StringCchPrintf(szBuf, 1024, __VA_ARGS__);\
	WriteConsole(GetStdHandle(STD_OUTPUT_HANDLE), szBuf, _tcslen(szBuf), NULL, NULL);

#define COM_CHECK_SUCCESS(hr, ...)\
	if(FAILED(hr))\
	{\
		COM_PRINTF(__VA_ARGS__);\
		goto __CLEAN_UP;\
	}
#define SAFE_RELEASE(I)\
	if((I) != NULL)\
	{\
		(I)->Release();\
		(I) = NULL;\
	}

#define DECLARE_OLEDB_INTERFACE(type) type* p##type = NULL

void ConnectSQLServerByDialog() //通过弹出对话框来链接SQL SERVER数据库
{
	DECLARE_BUFFER();
	DECLARE_OLEDB_INTERFACE(IDBPromptInitialize);
	DECLARE_OLEDB_INTERFACE(IDBInitialize);

	HWND hDesktop = GetDesktopWindow();
	HRESULT hRes = CoCreateInstance(CLSID_DataLinks, NULL, CLSCTX_INPROC_SERVER, IID_IDBPromptInitialize, (void**)&pIDBPromptInitialize);
	COM_CHECK_SUCCESS(hRes, _T("创建IDBPromptInitialize接口失败: %08x"), hRes);
	hRes = pIDBPromptInitialize->PromptDataSource(NULL, hDesktop, DBPROMPTOPTIONS_PROPERTYSHEET, 0, NULL, NULL, IID_IDBInitialize, (IUnknown**)&pIDBInitialize);
	COM_CHECK_SUCCESS(hRes, _T("弹出数据源对话框失败:%08x\n"), hRes);

	hRes = pIDBInitialize->Initialize();
	COM_CHECK_SUCCESS(hRes, _T("链接数据库失败:%08x\n"), hRes);
	COM_PRINTF(_T("链接数据库成功\n"));

	hRes = pIDBInitialize->Uninitialize();
__CLEAN_UP:
	SAFE_RELEASE(pIDBPromptInitialize);
	SAFE_RELEASE(pIDBInitialize);
	
}

void ConnectSQLServerByPropset() //通过属性链接SQL SERVER数据库
{
	DECLARE_BUFFER();
	DECLARE_OLEDB_INTERFACE(IDataInitialize);
	DECLARE_OLEDB_INTERFACE(IDBProperties);
	DECLARE_OLEDB_INTERFACE(IDBInitialize);
	CLSID clsid = {0};
	HRESULT hRes = CoCreateInstance(CLSID_MSDAINITIALIZE, NULL, CLSCTX_INPROC_SERVER, IID_IDataInitialize, (void**)&pIDataInitialize);
	COM_CHECK_SUCCESS(hRes, _T("创建接口IDBInitialize失败：%08x\n"), hRes);
	hRes = CLSIDFromProgID(_T("SQLOLEDB"), &clsid);
	COM_CHECK_SUCCESS(hRes, _T("查询SQLOLEDB CLSID 失败:%08x\n"), hRes);
	hRes = pIDataInitialize->CreateDBInstance(clsid, NULL,
		CLSCTX_INPROC_SERVER, NULL, IID_IDBInitialize,
		(IUnknown**)&pIDBInitialize);
	COM_CHECK_SUCCESS(hRes, _T("创建IDBInitialize接口失败:%08x\n"), hRes);

	//定义属性
	DBPROP dbProp[5] = {0};
	DBPROPSET dbPropsets[1] = {0};

	//设置数据库实例
	dbProp[0].dwPropertyID = DBPROP_INIT_DATASOURCE;
	dbProp[0].dwOptions = DBPROPOPTIONS_REQUIRED;
	dbProp[0].vValue.vt = VT_BSTR;
	dbProp[0].vValue.bstrVal = SysAllocString(OLESTR("LIU-PC\\SQLEXPRESS"));
	dbProp[0].colid = DB_NULLID;
	//设置数据库名称
	dbProp[1].dwPropertyID = DBPROP_INIT_CATALOG;
	dbProp[1].dwOptions = DBPROPOPTIONS_REQUIRED;
	dbProp[1].vValue.vt = VT_BSTR;
	dbProp[1].vValue.bstrVal = SysAllocString(OLESTR("Study"));
	dbProp[1].colid = DB_NULLID;
	//设置数据库用户名
	dbProp[2].dwPropertyID = DBPROP_AUTH_USERID;
	dbProp[2].vValue.vt = VT_BSTR;
	dbProp[2].vValue.bstrVal = SysAllocString(OLESTR("sa"));
	//设置数据库密码
	dbProp[3].dwPropertyID = DBPROP_AUTH_PASSWORD;
	dbProp[3].vValue.vt = VT_BSTR;
	dbProp[3].vValue.bstrVal = SysAllocString(OLESTR("123456"));
	
	dbPropsets[0].cProperties = 4;
	dbPropsets[0].guidPropertySet = DBPROPSET_DBINIT;
	dbPropsets[0].rgProperties = dbProp;
	//设置属性
	hRes = pIDBInitialize->QueryInterface(IID_IDBProperties, (void**)&pIDBProperties);
	COM_CHECK_SUCCESS(hRes, _T("查询IDBProperties接口失败:%08x\n"), hRes);
	hRes = pIDBProperties->SetProperties(1, dbPropsets);
	COM_CHECK_SUCCESS(hRes, _T("设置属性失败:%08x\n"), hRes);
	//链接数据库
	hRes = pIDBInitialize->Initialize();
	COM_CHECK_SUCCESS(hRes, _T("链接数据库失败:%08x\n"), hRes);
	COM_PRINTF(_T("链接数据库成功\n"));
	pIDBInitialize->Uninitialize();

__CLEAN_UP:
	SAFE_RELEASE(pIDataInitialize);
	SAFE_RELEASE(pIDBInitialize);
	SAFE_RELEASE(pIDBProperties);
}

void ConnectSQLServerByConnstr() //通过连接字符串连接数据库
{
	DECLARE_OLEDB_INTERFACE(IDataInitialize);
	DECLARE_OLEDB_INTERFACE(IDBInitialize);
	DECLARE_BUFFER();
	HRESULT hRes = CoCreateInstance(CLSID_MSDAINITIALIZE, NULL, CLSCTX_INPROC_SERVER, IID_IDataInitialize, (void**)&pIDataInitialize);
	COM_CHECK_SUCCESS(hRes, _T("创建IDataInitialize接口失败:%08x!\n"), hRes);

	//Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password = 123456;Initial Catalog=study;Data Source=LIU-PC\SQLEXPRESS
	hRes = pIDataInitialize->GetDataSource(NULL, CLSCTX_INPROC_SERVER, 
		OLESTR("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password = 123456;Initial Catalog=Study;Data Source=LIU-PC\\SQLEXPRESS;"), 
		IID_IDBInitialize, (IUnknown**)&pIDBInitialize);
	COM_CHECK_SUCCESS(hRes, _T("获取IDBInitialize接口失败:%08x!\n"), hRes);
	hRes = pIDBInitialize->Initialize();
	COM_CHECK_SUCCESS(hRes, _T("连接数据库失败:%08x!\n"), hRes);
	COM_PRINTF(_T("连接数据库成功\n"));
	pIDBInitialize->Uninitialize();
__CLEAN_UP:
	SAFE_RELEASE(pIDataInitialize);
	SAFE_RELEASE(pIDBInitialize);
}

void GetConnectString()
{
	DECLARE_OLEDB_INTERFACE(IDBPromptInitialize);
	DECLARE_OLEDB_INTERFACE(IDataInitialize);
	DECLARE_OLEDB_INTERFACE(IDBInitialize);
	DECLARE_BUFFER();
	
	LPOLESTR pConnStr = NULL;
	HWND hDeskTop = GetDesktopWindow();
	HRESULT hRes = CoCreateInstance(CLSID_DataLinks, NULL, CLSCTX_INPROC_SERVER, IID_IDBPromptInitialize, (void**)&pIDBPromptInitialize);
	COM_CHECK_SUCCESS(hRes, _T("创建IDBPromptInitialize接口失败:%08x!\n"), hRes);
	hRes = pIDBPromptInitialize->PromptDataSource(NULL, hDeskTop, DBPROMPTOPTIONS_PROPERTYSHEET, 0, NULL, NULL, IID_IDBInitialize, (IUnknown**)&pIDBInitialize);
	COM_CHECK_SUCCESS(hRes, _T("弹出数据源对话框失败:%08x\n"), hRes);


	hRes= pIDBPromptInitialize->QueryInterface(IID_IDataInitialize, (void**)&pIDataInitialize);
	COM_CHECK_SUCCESS(hRes, _T("创建IDataInitialize接口失败:%08x!\n"), hRes);

	hRes = pIDataInitialize->GetInitializationString(pIDBInitialize, TRUE, &pConnStr);
	COM_CHECK_SUCCESS(hRes, _T("获取连接字串失败失败:%08x\n"), hRes);

	COM_PRINTF(_T("连接字符串:%s"), pConnStr);
	SysAllocString(pConnStr);
__CLEAN_UP:
	SAFE_RELEASE(pIDataInitialize);
	SAFE_RELEASE(pIDBInitialize);
	SAFE_RELEASE(pIDBPromptInitialize);
}
int _tmain(int argc, TCHAR *argv[])
{
	CoInitialize(NULL);
	GetConnectString();
	CoUninitialize();
	return 0;
}
