
// opcDlg.cpp : ????
//

#include "stdafx.h"
#include "opc.h"
#include "opcDlg.h"
#include "afxdialogex.h"

#include "opcda.h"
#include "opccomn.h"
#include "opcerror.h"
#include "Item.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define NEXT_COUNT (100)


const CLSID CLSID_OPCServerList = {0x13486D51,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}};

CopcDlg::CopcDlg(CWnd* pParent /*=NULL*/)
    : CDialogEx(CopcDlg::IDD, pParent)
    , m_remoteAddress(_T(""))
    , m_dataName(_T(""))
    , m_dataType(_T(""))
    , m_dataQuality(_T(""))
    , m_dataValue(_T(""))
    , m_dataTime(_T(""))
{
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
    m_pIOPCItemMgt=nullptr;
    m_majorServer=nullptr;
    m_browser=nullptr;
    m_item=nullptr;
    m_callback=new MyOPCDataCallback();
    m_callback->AddRef();
}

void CopcDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_OPCSERVERLIST, m_ctlServerList);
    DDX_Text(pDX, IDC_REMOTESERVERADD, m_remoteAddress);
    DDX_Control(pDX, IDC_SERVERDETAILTREE, m_ctlServerTree);
    DDX_Control(pDX, IDC_DATACONTENT, m_ctlDataContent);
    DDX_Text(pDX, IDC_DATANAME, m_dataName);
    DDX_Text(pDX, IDC_DATATYPE, m_dataType);
    DDX_Text(pDX, IDC_DATAQUALITY, m_dataQuality);
    DDX_Text(pDX, IDC_DATAVALUE, m_dataValue);
    DDX_Text(pDX, IDC_DATATIME, m_dataTime);
}

BEGIN_MESSAGE_MAP(CopcDlg, CDialogEx)
    ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_CBN_SELCHANGE(IDC_OPCSERVERLIST, &CopcDlg::OnCbnSelchangeOpcserverlist)
    ON_BN_CLICKED(IDC_SELLOCAL, &CopcDlg::OnBnClickedSellocal)
    ON_BN_CLICKED(IDC_SELREMOTE, &CopcDlg::OnBnClickedSelremote)
    ON_BN_CLICKED(IDC_TRYTOCONNECTSERVER, &CopcDlg::OnBnClickedTrytoconnectserver)
    ON_BN_CLICKED(IDC_CONNECTION, &CopcDlg::OnBnClickedConnection)
    ON_NOTIFY(TVN_ITEMEXPANDING, IDC_SERVERDETAILTREE, &CopcDlg::OnTvnItemexpandingServerdetailtree)
    ON_NOTIFY(TVN_SELCHANGED, IDC_SERVERDETAILTREE, &CopcDlg::OnTvnSelchangedServerdetailtree)
    ON_LBN_DBLCLK(IDC_DATACONTENT, &CopcDlg::OnLbnDblclkDatacontent)
    ON_BN_CLICKED(IDC_DATAWRITE, &CopcDlg::OnBnClickedDatawrite)
END_MESSAGE_MAP()


BOOL CopcDlg::OnInitDialog()
{
    CDialogEx::OnInitDialog();

    SetIcon(m_hIcon, TRUE);         // ?????
    SetIcon(m_hIcon, FALSE);        // ?????

    CString szErrorText;
    HRESULT r1 = CoInitialize(NULL);
    if (r1 != S_OK)
    {
        if (r1 == S_FALSE)
        {
            MessageBox(_T("COM Library already initialized"),
                       _T("Error CoInitialize()"), MB_OK+MB_ICONEXCLAMATION);
        }
        else
        {
            szErrorText.Format(_T("Initialisation of COM Library failed. ErrorCode= %4x"), r1);
            MessageBox(szErrorText,_T("Error CoInitialize()"), MB_OK+MB_ICONERROR);
            SendMessage(WM_CLOSE);
            return TRUE;
        }
    }

    /*
    *   1. Choose local host as initial state
    */
    ((CButton*)GetDlgItem(IDC_SELLOCAL))->SetCheck(TRUE);
    GetDlgItem(IDC_REMOTESERVERADD)->EnableWindow(FALSE);
    m_srvSel=LOCALSERVER;

    m_callback->Register(this);


    return TRUE;  // ??????????,???? TRUE
}


void CopcDlg::OnPaint()
{
    if (IsIconic())
    {
        CPaintDC dc(this); // ??????????

        SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

        int cxIcon = GetSystemMetrics(SM_CXICON);
        int cyIcon = GetSystemMetrics(SM_CYICON);
        CRect rect;
        GetClientRect(&rect);
        int x = (rect.Width() - cxIcon + 1) / 2;
        int y = (rect.Height() - cyIcon + 1) / 2;


        dc.DrawIcon(x, y, m_hIcon);
    }
    else
    {
        CDialogEx::OnPaint();
    }
}


HCURSOR CopcDlg::OnQueryDragIcon()
{
    return static_cast<HCURSOR>(m_hIcon);
}

void CopcDlg::getOPCServer(CString _address)
{
    m_ctlServerList.ResetContent();
    CLSID cat_id=CATID_OPCDAServer20;
    IOPCServerList* server_list = 0;
    COSERVERINFO    si;
    MULTI_QI    qi;

    si.dwReserved1 = 0;
    si.dwReserved2 = 0;

    if(m_srvSel==REMOTESERVER&&_address.IsEmpty())
    {
        AfxMessageBox(_T("Please provide remote server's address!"));
        return;
    }

    if(_address && _address[0] != 0)
        si.pwszName = _address.GetBuffer(_address.GetLength());
    else
        si.pwszName = nullptr;
    si.pAuthInfo = nullptr;

    qi.pIID = &IID_IOPCServerList;
    qi.pItf = nullptr;
    qi.hr = 0;

    /*
    *   Connect to opc server
    */
    HRESULT hr = CoCreateInstanceEx(
                     CLSID_OPCServerList,
                     NULL,
                     CLSCTX_ALL,
                     &si,
                     1,
                     &qi);

    if(FAILED(hr) || FAILED(qi.hr))
    {
        CString msg(_T("Error connecting to OPC 2.0 Server Browser."));

        if( hr == REGDB_E_CLASSNOTREG )
        {
            CString msg(_T("Please install the OPC 2.0 Components."));
        }

        AfxMessageBox(msg);
    }
    else
    {
        server_list = (IOPCServerList*)qi.pItf;
        IEnumGUID* enum_guid = NULL;
        hr = server_list->EnumClassesOfCategories(
                 1,
                 &cat_id,
                 1,
                 &cat_id,
                 &enum_guid);
        if(SUCCEEDED(hr))
        {
            unsigned long count = 0;
            CLSID cls_id[NEXT_COUNT];
            do
            {
                hr = enum_guid->Next(NEXT_COUNT, cls_id, &count);
                for(unsigned int index = 0; index < count; index ++)
                {
                    LPOLESTR prog_id;
                    LPOLESTR user_type;
                    HRESULT hr2 = server_list->GetClassDetails(cls_id[index], &prog_id, &user_type);
                    /*
                    *   Save the CLSID for further connection;
                    */
                    if(SUCCEEDED(hr2))
                    {
                        CString name;
                        name.Format(_T("%s,%s"),(LPCTSTR)prog_id,(LPCTSTR)user_type);
                        int pos=m_ctlServerList.AddString(name);
                        m_ctlServerList.SetItemData(pos,DWORD(new CLSID(cls_id[index])));
                        CoTaskMemFree(prog_id);
                        CoTaskMemFree(user_type);
                    }
                }
            }
            while(hr == S_OK);

            enum_guid->Release();
            server_list->Release();
        }
    }
}

HRESULT CopcDlg::connect2CurrentServer(void)
{
    /*
    *   Local Server
    */

    HRESULT hr=S_OK;
    if(m_srvSel==LOCALSERVER)
    {

        hr = CoCreateInstance(
                 *m_majorServer,
                 NULL,
                 CLSCTX_LOCAL_SERVER,
                 IID_IOPCServer,
                 (LPVOID*)&m_opcService);

        if(FAILED(hr) || m_opcService == nullptr)     //????
        {

            AfxMessageBox(_T("Connect failed!"));
            return hr;
        }
    }
    else
    {
        /*
        *   Remote server
        */
        COSERVERINFO    si;
        MULTI_QI    qi;

        si.dwReserved1 = 0;
        si.pAuthInfo = nullptr;
        si.pwszName = m_remoteAddress.GetBuffer(m_remoteAddress.GetLength());
        si.dwReserved2 = 0;

        qi.pIID = &IID_IOPCServer;
        qi.pItf = nullptr;
        qi.hr = 0;

        hr = CoCreateInstanceEx(
                 *m_majorServer,
                 nullptr,
                 CLSCTX_REMOTE_SERVER,
                 &si,
                 1,
                 &qi);
        if(FAILED(hr))
        {
            AfxMessageBox(_T("Connect failed!"));

            return hr;
        }
        if(FAILED(qi.hr) || (qi.pItf == NULL))
        {
            AfxMessageBox(_T("MULTI_QI failed: "));


            return hr;
        }

        m_opcService = (IOPCServer*)qi.pItf;
    }

    return hr;
}

HRESULT CopcDlg::addGroup(void)
{
    LONG        TimeBias = 0;
    FLOAT       PercentDeadband = 0.0;
    DWORD       RevisedUpdateRate;
    HRESULT hr;
    /*
    *   Add a group
    */
    hr=m_opcService->AddGroup(L"Amos",     // [in] group name
                              TRUE,              // [in] not active, so no OnDataChange callback
                              500,                // [in] Request this Update Rate from Server
                              1,                  // [in] Client Handle, not necessary in this sample
                              &TimeBias,          // [in] No time interval to system UTC time
                              &PercentDeadband,   // [in] No Deadband, so all data changes are reported
                              LOCALE_ID,          // [in] Server uses english language to for text values
                              &m_GrpSrvHandle,    // [out] Server handle to identify this group in later calls
                              &RevisedUpdateRate, // [out] The answer from Server to the requested Update Rate
                              IID_IOPCGroupStateMgt,    // [in] requested interface type of the group object
                              (LPUNKNOWN*)&m_IOPCGroupMgt); // [out] pointer to the requested interface

    if (FAILED(hr))
    {
        MessageBox(_T("Can't add Group to Server!"),
                   _T("Error AddGroup()"), MB_OK+MB_ICONERROR);
        m_opcService->Release();
        m_opcService = nullptr;
        return hr;
    }

    /*
    *   Once connect successfully, Update tree control.
    *   This is UI operation
    */
    OPCNAMESPACETYPE nameSpaceType;

    m_opcService->QueryInterface(IID_IOPCBrowseServerAddressSpace,(void**)&m_browser);
    m_browser->QueryOrganization(&nameSpaceType);
    VARTYPE vtype=VT_EMPTY;
    if(nameSpaceType==OPC_NS_HIERARCHIAL)
    {
        HTREEITEM root = m_ctlServerTree.InsertItem(_T("root"));
        m_ctlServerTree.SetItemData(root, 1);
        IEnumString* enum_string = NULL;

        hr = m_browser->BrowseOPCItemIDs(
                 OPC_BRANCH,
                 L"*",
                 vtype,
                 0,
                 &enum_string);

        if(SUCCEEDED(hr))
        {
            LPWSTR  name[NEXT_COUNT];
            ULONG   count = 0;

            do
            {
                hr = enum_string->Next(
                         NEXT_COUNT,
                         &name[0],
                         &count);

                for(ULONG index = 0; index < count; index ++)
                {
                    CString item_name(name[index]);
                    HTREEITEM item = m_ctlServerTree.InsertItem(item_name, root);
                    m_ctlServerTree.SetItemData(item, 0);
                    m_ctlServerTree.InsertItem(_T("Dummy"), item);
                    CoTaskMemFree(name[index]);
                }
            }
            while(hr == S_OK);
            enum_string->Release();
        }
        m_ctlServerTree.SelectItem(root);
        m_ctlServerTree.Expand(root,TVE_EXPAND);
    }
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    return hr;
}

CItem* CopcDlg::addItem(CString _name)
{
    HRESULT hr=S_OK;
    OPCITEMRESULT* results;
    HRESULT* errors;
    CItem* item=new CItem();
    item->name=_name;

    OPCITEMDEF  item_def;

    /*
    *   Get full name from server. !!Very Important
    */
    LPWSTR iname=nullptr;

    hr=m_browser->GetItemID(T2OLE(_name.GetBuffer(_name.GetLength())),&iname);

    item_def.szItemID = iname;
    item_def.szAccessPath = _T("");
    item_def.pBlob = nullptr;
    item_def.dwBlobSize = 0;
    item_def.bActive = TRUE;
    item_def.hClient = (OPCHANDLE)item;
    item_def.vtRequestedDataType = VT_EMPTY;

    hr = m_IOPCGroupMgt->QueryInterface(IID_IOPCItemMgt, (LPVOID*)&m_pIOPCItemMgt);
    if(FAILED(hr))
    {
        delete item;
        return nullptr;
    }
    hr = m_pIOPCItemMgt->AddItems(1, &item_def, &results, &errors);
    if(FAILED(hr))
    {
        delete item;
        return nullptr;
    }
    item->hServer=results->hServer;
    item->nativeType=results->vtCanonicalDataType;

    /*
    *  Add callback to monitor the data change
    */
    setCallback();

    /*
    *   Read initial value;
    */

    read(item);

    return item;
}

void CopcDlg::setCallback()
{

    AtlAdvise(m_IOPCGroupMgt,m_callback,
              IID_IOPCDataCallback,
              &m_callbackConnection);
}

HRESULT CopcDlg::read(CItem* _item)
{
    HRESULT hr;
    IOPCSyncIO* sIO;
    HRESULT* errors;
    hr=m_IOPCGroupMgt->QueryInterface(IID_IOPCSyncIO,(void**)&sIO);
    if(SUCCEEDED(hr))
    {
        OPCITEMSTATE* item_state = NULL;
        hr = sIO->Read(
                 OPC_DS_DEVICE,
                 1,
                 &_item->hServer,
                 &item_state,
                 &errors);
        if(SUCCEEDED(hr))
        {
            ASSERT(item_state->hClient == (OPCHANDLE)_item);
            _item->quality = item_state->wQuality;
            _item->value = item_state->vDataValue;
            _item->timeStamp=item_state->ftTimeStamp;

            VariantClear(&item_state->vDataValue);
            CoTaskMemFree(item_state);
            CoTaskMemFree(errors);
        }
        else
        {
            AfxMessageBox(_T("Read Failed!"));
            return hr;
        }

        sIO->Release();
    }

    return hr;
}

HRESULT CopcDlg::write(CItem* _item)
{
    HRESULT hr;
    IOPCSyncIO* sio;
    HRESULT* errors = NULL;
    CString str;
    COleVariant value;

    /*
    *   Get value from m_dataValue;
    */
    UpdateData();
    Str2Variant(m_dataValue,m_item->nativeType,value);

    hr=m_IOPCGroupMgt->QueryInterface(IID_IOPCSyncIO,(void**)&sio);
    if(SUCCEEDED(hr))
    {
        HRESULT hr = sio->Write(
                         1,
                         &_item->hServer,
                         &value,
                         &errors);
        if(SUCCEEDED(hr))
        {
            if(FAILED(errors[0]))
            {
                str.Format(_T("%s"),errors[0]);
                AfxMessageBox(str);
            }
            CoTaskMemFree(errors);
        }
        sio->Release();
    }
    return hr;
}

void CopcDlg::UpdateItem(OPCHANDLE _item, VARIANT *_value,WORD _quality, FILETIME _time)
{
    if(_item==(OPCHANDLE)m_item)
    {
        m_item->quality = _quality;
        VariantCopy(&(m_item->value),_value);
        m_item->timeStamp=_time;
    }
    updateView(m_item);
}

void CopcDlg::updateView(CItem* _item)
{
    switch(_item->quality)
    {
        case OPC_QUALITY_GOOD:
        {
            m_dataQuality=_T("Good");
        }
        break;
        case OPC_QUALITY_BAD:
        {
            m_dataQuality=_T("Bad");
        }
        break;
        default:
        {
            m_dataQuality=_T("Unknow");
        }
        break;
    }

    switch(_item->nativeType)
    {
        case VT_I2:
        {
            m_dataType = _T("int");
            m_dataValue.Format(_T("%d"),_item->value.iVal);
        }
        break;
        case VT_UI1:
        {
            m_dataType = _T("uint");
            m_dataValue.Format(_T("%u"),_item->value.ulVal);
        }
        break;
        case VT_I4:
        {
            m_dataType = _T("long");
            m_dataValue.Format(_T("%lld"),_item->value.lVal);
        }
        break;
        case VT_UI2:
        {
            m_dataType = _T("word");
            m_dataValue.Format(_T("%d"), _item->value.uiVal);
        }
        break;
        case VT_UI4:
        {
            m_dataType = _T("Dword");
            m_dataValue.Format(_T("%d"), _item->value.ulVal);
        }
        break;
        case VT_R4:
        {
            m_dataType = _T("float");
            m_dataValue.Format(_T("%f"),_item->value.fltVal);
        }
        break;
        case VT_R8:
        {
            m_dataType = _T("double");
            m_dataValue.Format(_T("%lld"),_item->value.dblVal);
        }
        break;
        case VT_BOOL:
        {
            m_dataType = _T("bool");
            m_dataValue=_item->value.boolVal?_T("TRUE"):_T("FALSE");
        }
        break;
        case VT_BSTR:
        {
            m_dataType = _T("string");
            m_dataValue.Format(_T("%s"),_item->value.bstrVal);
        }
        break;
        default:
            m_dataType = _T("Unknown");
            m_dataValue=_T("");
            break;
    }
    CTime time(_item->timeStamp);
    m_dataTime.Format(_T("%s"),time.Format("%c"));
    UpdateData(FALSE);
}

HRESULT CopcDlg::surfTree(HTREEITEM _item)
{
    //????
    HRESULT hr = S_OK;
    if(_item != NULL)
    {
        HTREEITEM parent = m_ctlServerTree.GetParentItem(_item);
        hr = surfTree(parent);
        if(SUCCEEDED(hr))
        {
            CString node(m_ctlServerTree.GetItemText(_item));
            if(node != _T("root"))
            {
                hr = m_browser->ChangeBrowsePosition(
                         OPC_BROWSE_DOWN,
                         T2OLE(node.GetBuffer(node.GetLength())));
            }
        }
    }
    return hr;
}


void CopcDlg::OnCbnSelchangeOpcserverlist()
{
    // TODO: ??????????????
    int sel=m_ctlServerList.GetCurSel();
    m_majorServer=(CLSID*)m_ctlServerList.GetItemData(sel);
}

void CopcDlg::OnBnClickedSellocal()
{
    // TODO: ??????????????
    GetDlgItem(IDC_REMOTESERVERADD)->EnableWindow(FALSE);
    m_srvSel=LOCALSERVER;
}


void CopcDlg::OnBnClickedSelremote()
{
    // TODO: ??????????????
    GetDlgItem(IDC_REMOTESERVERADD)->EnableWindow(TRUE);
    m_srvSel=REMOTESERVER;
}


void CopcDlg::OnBnClickedTrytoconnectserver()
{
    // TODO: ??????????????
    UpdateData();
    if(m_srvSel==LOCALSERVER)
    {
        getOPCServer();
    }
    else
    {
        getOPCServer(m_remoteAddress);
    }
}


void CopcDlg::OnBnClickedConnection()
{
    /*
    *   Local Server
    */
    m_ctlServerTree.DeleteAllItems();

    if(!SUCCEEDED(connect2CurrentServer()))
        return;

    addGroup();

}

void CopcDlg::OnTvnItemexpandingServerdetailtree(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
    // TODO: ??????????????

    HTREEITEM start = pNMTreeView->itemNew.hItem;
    if(!start)
        return;

    DWORD expand = m_ctlServerTree.GetItemData(start);
    if(!expand)
    {
        m_ctlServerTree.SetItemData(start, 1);

        HTREEITEM item = m_ctlServerTree.GetNextItem(start, TVGN_CHILD);
        while(item)
        {
            m_ctlServerTree.DeleteItem(item);
            item = m_ctlServerTree.GetNextItem(start, TVGN_CHILD);
        }


        HRESULT hr = surfTree(start);
        if(SUCCEEDED(hr))
        {
            IEnumString* enum_string = NULL;
            hr = m_browser->BrowseOPCItemIDs(
                     OPC_BRANCH,
                     L"*",
                     VT_EMPTY,
                     0,
                     &enum_string);
            if(SUCCEEDED(hr))
            {
                LPWSTR  name[NEXT_COUNT];
                ULONG   count = 0;
                do
                {
                    hr = enum_string->Next(NEXT_COUNT, &name[0], &count);
                    for(ULONG index = 0; index < count; index++)
                    {
                        CString item_name(name[index]);
                        HTREEITEM item = m_ctlServerTree.InsertItem(item_name, start);
                        m_ctlServerTree.SetItemData(item, 0);
                        m_ctlServerTree.InsertItem(_T("Dummy"), item);
                        CoTaskMemFree(name[index]);
                    }
                }
                while(hr == S_OK);
                enum_string->Release();
            }
        }

        hr = S_OK;
        for(int i = 0; i < 16 && SUCCEEDED(hr); i++)
        {
            hr = m_browser->ChangeBrowsePosition(
                     OPC_BROWSE_UP,
                     L"");
        }
    }
    *pResult = 0;
}

void CopcDlg::OnTvnSelchangedServerdetailtree(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
    // TODO: ??????????????
    HTREEITEM item = m_ctlServerTree.GetSelectedItem();
    HRESULT hr = surfTree(item);
    if(SUCCEEDED(hr))
    {
        m_ctlDataContent.ResetContent();

        IEnumString* enum_string = 0;
        /*
        *   Surf the leaf node
        */
        hr = m_browser->BrowseOPCItemIDs(
                 OPC_LEAF,
                 _T(""),
                 VT_EMPTY,
                 0,
                 &enum_string);
        if(SUCCEEDED(hr))
        {
            LPWSTR name[NEXT_COUNT];
            ULONG count = 0;
            do
            {
                hr = enum_string->Next(NEXT_COUNT, &name[0], &count);
                for(ULONG index = 0; index < count; index++)
                {
                    CString tag_name(name[index]);
                    m_ctlDataContent.AddString(tag_name);

                    CoTaskMemFree(name[index]);
                }

            }
            while(hr == S_OK);
            enum_string->Release();
        }
    }

    hr = S_OK;
    for(int i = 0; i < 16 && SUCCEEDED(hr); i++)
    {
        hr = m_browser->ChangeBrowsePosition(
                 OPC_BROWSE_UP,
                 L"");
    }
    *pResult = 0;
}


void CopcDlg::OnLbnDblclkDatacontent()
{
    // TODO: ??????????????
    HTREEITEM treeItem=m_ctlServerTree.GetSelectedItem();
    HRESULT hr=surfTree(treeItem);  // Let the browser get into this position. !!!!!!
    if(FAILED(hr))
    {
        AfxMessageBox(_T("Add failed!"));
        return;
    }
    int sel=m_ctlDataContent.GetCurSel();
    if(sel!=LB_ERR)
    {
        m_ctlDataContent.GetText(sel,m_dataName);
    }

    LPWSTR item_fullName=nullptr;
    hr = m_browser->GetItemID(T2OLE(m_dataName.GetBuffer(0)), &item_fullName);
    m_dataName.Format(_T("%s"),item_fullName);
    CItem* item=addItem(m_dataName);
    if(item)
    {
        updateView(item);
        if(m_item!=nullptr)
        {
            delete m_item;
            m_item=nullptr;
        }
        m_item=item;
    }

    for(int i = 0; i < 16 && SUCCEEDED(hr); i++)
    {
        hr = m_browser->ChangeBrowsePosition(
                 OPC_BROWSE_UP,
                 L"");
    }

    UpdateData(FALSE);

}

void CopcDlg::OnBnClickedDatawrite()
{
    // TODO: ??????????????
    write(m_item);
}

BOOL CopcDlg::Str2Variant(CString _valuestr,unsigned short _vtype, VARIANT& _itmV)
{
    CT2A pBuff(_valuestr);
    _itmV.vt = _vtype;
    switch(_itmV.vt )
    {
        case VT_I1:
        case VT_UI1:        // BYTE
            _itmV.bVal =(BYTE)atoi(pBuff);
            break;
        case VT_I2:         // SHORT
            _itmV.iVal=atoi(pBuff);
            break;
        case VT_UI2:        // UNSIGNED SHORT
            _itmV.uiVal =(USHORT)atoi(pBuff);
            break;
        case VT_I4:         // LONG
            _itmV.lVal= atol(pBuff);
            break;
        case VT_UI4:        // UNSIGNED LONG
            _itmV.ulVal =(ULONG)atol(pBuff);
            break;
        case VT_INT:        // INTEGER
            _itmV.intVal =atoi(pBuff);
            break;
        case VT_UINT:       // UNSIGNED INTEGER
            _itmV.uintVal =( USHORT)atoi(pBuff);
            break;
        case VT_R4:         // FLOAT
            _itmV.fltVal =(float)atof(pBuff);
            break;
        case VT_R8:         // DOUBLE
            _itmV.dblVal =atof(pBuff);
            break;
        case VT_BSTR:       //BSTR
            VariantClear((VARIANT *)&_itmV.bstrVal);//SysAllocString((BSTR)
            //itmV.bstrVal =_com_util::ConvertStringToBSTR(pBuff);
            _itmV.bstrVal = _bstr_t(pBuff);
            _itmV.vt = VT_BSTR;
            break;
        case VT_BOOL:
            if(!stricmp("TRUE",pBuff))
                _itmV.boolVal=TRUE;
            else
                _itmV.boolVal=FALSE;
            break;
        case VT_DATE:
            //  VariantTimeToSystemTime (itmV.date, &st);
            //  ct = st;
            //  cs = ct.Format ("%H:%M:%S %d %b %Y");
            //  strncpy (buf, cs, 200);
            //  return;
            break;
        default:
            return false;
    }
    return true;
}