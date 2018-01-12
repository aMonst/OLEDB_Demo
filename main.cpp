#include <tchar.h>
#include <windows.h>
#include <strsafe.h>

#define COM_NO_WINDOWS_H    //如果已经包含了Windows.h或不使用其他Windows库函数时
#define OLEDBVER 0x0260     //MSDAC2.6版
#include <oledb.h>
#include <oledberr.h>

#define GRS_ALLOC(sz)		HeapAlloc(GetProcessHeap(),0,sz)
#define GRS_CALLOC(sz)		HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY,sz)
#define GRS_SAFEFREE(p)		if(NULL != p){HeapFree(GetProcessHeap(),0,p);p=NULL;}

#define GRS_USEPRINTF() TCHAR pBuf[1024] = {}
#define GRS_PRINTF(...) \
	GRS_USEPRINTF();\
	StringCchPrintf(pBuf,1024,__VA_ARGS__);\
	WriteConsole(GetStdHandle(STD_OUTPUT_HANDLE),pBuf,lstrlen(pBuf),NULL,NULL);

#define GRS_SAFERELEASE(I)\
	if(NULL != (I))\
	{\
		(I)->Release();\
		(I)=NULL;\
	}

#define GRS_COM_CHECK(hr,...)\
	if(FAILED(hr))\
	{\
		GRS_PRINTF(__VA_ARGS__);\
		goto CLEAR_UP;\
	}

int _tmain(int argc, TCHAR* argv[])
{
	CoInitialize(NULL);
	//创建OLEDB init接口
	IDBInitialize *pDBInit = NULL;
	IDBProperties *pIDBProperties = NULL;
	//设置链接属性
	DBPROPSET dbPropset[1] = {0};
	DBPROP dbProps[5] = {0};
	CLSID clsid_MSDASQL = {0}; //sql server 的数据源对象
	
	HRESULT hRes = CLSIDFromProgID(_T("SQLOLEDB"), &clsid_MSDASQL);
	GRS_COM_CHECK(hRes, _T("获取SQLOLEDB的CLSID失败，错误码：0x%08x\n"), hRes);
	hRes = CoCreateInstance(clsid_MSDASQL, NULL, CLSCTX_INPROC_SERVER, IID_IDBInitialize,(void**)&pDBInit);
	GRS_COM_CHECK(hRes, _T("无法创建IDBInitialize接口，错误码：0x%08x\n"), hRes);

	//指定数据库实例名，这里使用了别名local，指定本地默认实例
	dbProps[0].dwPropertyID = DBPROP_INIT_DATASOURCE;
	dbProps[0].dwOptions = DBPROPOPTIONS_REQUIRED;
	dbProps[0].vValue.vt = VT_BSTR;
	dbProps[0].vValue.bstrVal = SysAllocString(OLESTR("LIU-PC\\SQLEXPRESS"));
	dbProps[0].colid = DB_NULLID;

	//指定数据库库名
	dbProps[1].dwPropertyID = DBPROP_INIT_CATALOG;
	dbProps[1].dwOptions = DBPROPOPTIONS_REQUIRED;
	dbProps[1].vValue.vt = VT_BSTR;
	dbProps[1].vValue.bstrVal = SysAllocString(OLESTR("Study"));
	dbProps[1].colid = DB_NULLID;

	//指定链接数据库的用户名
	dbProps[2].dwPropertyID = DBPROP_AUTH_USERID;
	dbProps[2].vValue.vt = VT_BSTR;
	dbProps[2].vValue.bstrVal = SysAllocString(OLESTR("sa"));
	
	//指定链接数据库的用户密码
	dbProps[3].dwPropertyID = DBPROP_AUTH_PASSWORD;
	dbProps[3].vValue.vt = VT_BSTR;
	dbProps[3].vValue.bstrVal = SysAllocString(OLESTR("123456"));
	
	
	//设置属性
	hRes = pDBInit->QueryInterface(IID_IDBProperties, (void**)&pIDBProperties);
	GRS_COM_CHECK(hRes, _T("查询IDBProperties接口失败, 错误码:%08x\n"), hRes);
	dbPropset->guidPropertySet = DBPROPSET_DBINIT;
	dbPropset[0].cProperties = 4;
	dbPropset[0].rgProperties = dbProps;
	hRes = pIDBProperties->SetProperties(1, dbPropset);
	GRS_COM_CHECK(hRes, _T("设置属性失败, 错误码:%08x\n"), hRes);

	//链接数据库
	hRes = pDBInit->Initialize();
	GRS_COM_CHECK(hRes, _T("链接数据库失败：错误码:%08x\n"), hRes);
	//do something
	pDBInit->Uninitialize();

	GRS_PRINTF(_T("数据库操作成功!!!!!\n"));
CLEAR_UP:
	GRS_SAFEFREE(pDBInit);
	GRS_SAFEFREE(pIDBProperties);
	CoUninitialize();
	return 0;
}
