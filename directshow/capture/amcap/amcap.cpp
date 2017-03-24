//------------------------------------------------------------------------------
// File: AMCap.cpp
//
// Desc: Audio/Video Capture sample for DirectShow
//
//
// Copyright (c) Microsoft Corporation.  All rights reserved.
//------------------------------------------------------------------------------

#include "stdafx.h"
#include <dbt.h>
#include <mmreg.h>
#include <msacm.h>
#include <fcntl.h>
#include <io.h>
#include <stdio.h>
#include <commdlg.h>
#include <strsafe.h>
#include "stdafx.h"
#include "amcap.h"
#include "status.h"
#include "SampleCGB.h"
#include "sample-grabber.hpp"

#define check(expr) if (!SUCCEEDED(hr = (expr))) goto fail

class CallbackObject : public ISampleGrabberCB{
public:

	CallbackObject() {};

	STDMETHODIMP QueryInterface(REFIID riid, void **ppv)
	{
		if (NULL == ppv) return E_POINTER;
		if (riid == __uuidof(IUnknown)) {
			*ppv = static_cast<IUnknown*>(this);
			return S_OK;
		}
		if (riid == __uuidof(ISampleGrabberCB)) {
			*ppv = static_cast<ISampleGrabberCB*>(this);
			return S_OK;
		}
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() { return 42; }
	STDMETHODIMP_(ULONG) Release() { return 42; }

	//ISampleGrabberCB
	STDMETHODIMP SampleCB(double SampleTime, IMediaSample *pSample)
	{
		printf("sample-cb tm %f %d\n", SampleTime, (int)pSample->GetSize());
		return S_OK;
	}
	STDMETHODIMP BufferCB(double SampleTime, BYTE *pBuffer, long BufferLen)
	{
		printf("buffer-cb tm %f %d\n", SampleTime, (int)BufferLen);
		return S_OK;
	}
};

//------------------------------------------------------------------------------
// Macros
//------------------------------------------------------------------------------
#define ABS(x) (((x) > 0) ? (x) : -(x))

// An application can advertise the existence of its filter graph
// by registering the graph with a global Running Object Table (ROT).
// The GraphEdit application can detect and remotely view the running
// filter graph, allowing you to 'spy' on the graph with GraphEdit.
//
// To enable registration in this sample, define REGISTER_FILTERGRAPH.
//

// {C1F400A0-3F08-11D3-9F0B-006008039E37}
DEFINE_GUID(CLSID_SampleGrabber,
	0xC1F400A0, 0x3F08, 0x11D3, 0x9F, 0x0B, 0x00, 0x60, 0x08, 0x03, 0x9E, 0x37); //qedit.dll

//------------------------------------------------------------------------------
// Global data
//------------------------------------------------------------------------------
static HINSTANCE ghInstApp=0;
static HACCEL ghAccel=0;
static HFONT  ghfontApp=0;
static TEXTMETRIC gtm={0};
static TCHAR gszAppName[]=TEXT("AMCAP");
static HWND ghwndApp=0, ghwndStatus=0;
static HDEVNOTIFY ghDevNotify=0;
static PUnregisterDeviceNotification gpUnregisterDeviceNotification=0;
static PRegisterDeviceNotification gpRegisterDeviceNotification=0;
static DWORD g_dwGraphRegister=0;

struct _capstuff
{
	ISampleGrabber *pGrabber;
    ISampleCaptureGraphBuilder *pBuilder;
	IBaseFilter *pVW;
    IMediaEventEx *pME;
    IAMDroppedFrames *pDF;
    IAMVideoCompression *pVC;
    IAMVfwCaptureDialogs *pDlg;
    IAMStreamConfig *pVSC;      // for video cap
    IBaseFilter *pRender;
    IBaseFilter *pVCap;
    IGraphBuilder *pFg;
    IFileSinkFilter *pSink;
	CallbackObject callback;
    int  iMasterStream;
    BOOL fCaptureGraphBuilt;
    BOOL fPreviewGraphBuilt;
	BOOL previewing;
    bool fDeviceMenuPopulated;
    IMoniker *rgpmVideoMenu[10];
    IMoniker *pmVideo;
    WCHAR wachFriendlyName[120];
    int iFormatDialogPos;
    int iSourceDialogPos;
    int iDisplayDialogPos;
    int iVCapDialogPos;
    int iVCapCapturePinDialogPos;
    int iVCapPreviewPinDialogPos;
    long lDroppedBase;
    long lNotBase;
    BOOL fPreviewFaked;
    int iVideoInputMenuPos;
    LONG NumberOfVideoInputs;
    HMENU hMenuPopup;
    int iNumVCapDevices;
};

// can't thread local due to COM object apartments
// XXX TODO convert to object
static struct _capstuff gcap;

//------------------------------------------------------------------------------
// Function Prototypes
//------------------------------------------------------------------------------
typedef LONG(PASCAL *LPWNDPROC)(HWND, UINT, WPARAM, LPARAM); // pointer to a window procedure
static LONG WINAPI AppWndProc(HWND hwnd, UINT uiMessage, WPARAM wParam, LPARAM lParam);
static LONG PASCAL AppCommand(HWND hwnd, unsigned msg, WPARAM wParam, LPARAM lParam);
static BOOL CALLBACK AboutDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

static void ErrMsg(LPTSTR sz,...);

static void TearDownGraph(void);
static BOOL BuildPreviewGraph();

static void AddDevicesToMenu();
static void ChooseDevices(TCHAR *szVideo);
static void ChooseDevices(IMoniker *pmVideo);

static BOOL InitCapFilters();
static void FreeCapFilters();

static BOOL StopPreview();
static BOOL StartPreview();

static void MakeMenuOptions();
static void OnClose();



//------------------------------------------------------------------------------
// Name: SetAppCaption()
// Desc: Set the caption to be the application name followed by the capture file
//------------------------------------------------------------------------------
static void SetAppCaption()
{
    TCHAR tach[_MAX_PATH + 80];

    StringCchCopy(tach, NUMELMS(tach), gszAppName);

    SetWindowText(ghwndApp, tach);
}


/*----------------------------------------------------------------------------*\
|   AppInit( hInst, hPrev)                                                     |
|                                                                              |
|   Description:                                                               |
|       This is called when the application is first loaded into               |
|       memory.  It performs all initialization that doesn't need to be done   |
|       once per instance.                                                     |
|                                                                              |
|   Arguments:                                                                 |
|       hInstance       instance handle of current instance                    |
|       hPrev           instance handle of previous instance                   |
|                                                                              |
|   Returns:                                                                   |
|       TRUE if successful, FALSE if not                                       |
|                                                                              |
\*----------------------------------------------------------------------------*/
static BOOL AppInit(HINSTANCE hInst, HINSTANCE hPrev, int sw)
{
    WNDCLASS    cls;
    HDC         hdc;

    const DWORD  dwExStyle = 0;

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    /* Save instance handle for DialogBoxs */
    ghInstApp = hInst;

    ghAccel = LoadAccelerators(hInst, MAKEINTATOM(ID_APP));

    if(!hPrev)
    {
        /*
        *  Register a class for the main application window
        */
        cls.hCursor        = LoadCursor(NULL,IDC_ARROW);
        cls.hIcon          = NULL;
        cls.lpszMenuName   = MAKEINTATOM(ID_APP);
        cls.lpszClassName  = MAKEINTATOM(ID_APP);
        cls.hbrBackground  = (HBRUSH)(COLOR_WINDOW + 1);
        cls.hInstance      = hInst;
        cls.style          = CS_BYTEALIGNCLIENT | CS_VREDRAW | CS_HREDRAW | CS_DBLCLKS;
        cls.lpfnWndProc    = (WNDPROC) AppWndProc;
        cls.cbWndExtra     = 0;
        cls.cbClsExtra     = 0;

        if(!RegisterClass(&cls))
            return FALSE;
    }

    // Is this necessary?
    ghfontApp = (HFONT)GetStockObject(ANSI_VAR_FONT);
    hdc = GetDC(NULL);
    SelectObject(hdc, ghfontApp);
    GetTextMetrics(hdc, &gtm);
    ReleaseDC(NULL, hdc);

    ghwndApp=CreateWindowEx(dwExStyle,
                            MAKEINTATOM(ID_APP),    // Class name
                            gszAppName,             // Caption
                            // Style bits
                            WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                            CW_USEDEFAULT, 0,       // Position
                            320,300,                // Size
                            (HWND)NULL,             // Parent window (no parent)
                            (HMENU)NULL,            // use class menu
                            hInst,                  // handle to window instance
                            (LPSTR)NULL             // no params to pass on
                            );

    // create the status bar
    statusInit(hInst, hPrev);
    ghwndStatus = CreateWindowEx(0,
                            szStatusClass,
                            NULL,
                            WS_CHILD|WS_BORDER|WS_VISIBLE|WS_CLIPSIBLINGS,
                            0, 0,
                            0, 0,
                            ghwndApp,
                            NULL,
                            hInst,
                            NULL);
    if(ghwndStatus == NULL)
        return(FALSE);

    ShowWindow(ghwndApp,sw);

    // Read the list of devices to use from win.ini
    ZeroMemory(gcap.rgpmVideoMenu, sizeof(gcap.rgpmVideoMenu));
    gcap.pmVideo = 0;

	gcap.previewing = 0;


    gcap.fDeviceMenuPopulated = false;
    AddDevicesToMenu();

    // Instantiate the capture filters we need to do the menu items.
    // This will start previewing, if desired
    //
    // Make these the official devices we're using

    ChooseDevices(L"");

    // Register for device add/remove notifications
    DEV_BROADCAST_DEVICEINTERFACE filterData;
    ZeroMemory(&filterData, sizeof(DEV_BROADCAST_DEVICEINTERFACE));

    filterData.dbcc_size = sizeof(DEV_BROADCAST_DEVICEINTERFACE);
    filterData.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
    filterData.dbcc_classguid = AM_KSCATEGORY_CAPTURE;

    gpUnregisterDeviceNotification = NULL;
    gpRegisterDeviceNotification = NULL;
    // dynload device removal APIs
    {
        HMODULE hmodUser = GetModuleHandle(TEXT("user32.dll"));
        ASSERT(hmodUser);       // we link to user32
        gpUnregisterDeviceNotification = (PUnregisterDeviceNotification)
        GetProcAddress(hmodUser, "UnregisterDeviceNotification");

        // m_pRegisterDeviceNotification is prototyped differently in unicode
        gpRegisterDeviceNotification = (PRegisterDeviceNotification)
            GetProcAddress(hmodUser,
            "RegisterDeviceNotificationW"
        );

        // failures expected on older platforms.
        ASSERT(gpRegisterDeviceNotification && gpUnregisterDeviceNotification ||
              !gpRegisterDeviceNotification && !gpUnregisterDeviceNotification);
    }

    ghDevNotify = NULL;

    if(gpRegisterDeviceNotification)
    {
        ghDevNotify = gpRegisterDeviceNotification(ghwndApp, &filterData, DEVICE_NOTIFY_WINDOW_HANDLE);
        ASSERT(ghDevNotify != NULL);
    }

    SetAppCaption();
    return TRUE;
}

static void IMonRelease(IMoniker *&pm)
{
    if(pm)
    {
        pm->Release();
        pm = 0;
    }
}

/*----------------------------------------------------------------------------*\
|   WinMain( hInst, hPrev, lpszCmdLine, cmdShow )                              |
|                                                                              |
|   Description:                                                               |
|       The main procedure for the App.  After initializing, it just goes      |
|       into a message-processing loop until it gets a WM_QUIT message         |
|       (meaning the app was closed).                                          |
|                                                                              |
|   Arguments:                                                                 |
|       hInst           instance handle of this instance of the app            |
|       hPrev           instance handle of previous instance, NULL if first    |
|       szCmdLine       ->null-terminated command line                         |
|       cmdShow         specifies how the window is initially displayed        |
|                                                                              |
|   Returns:                                                                   |
|       The exit code as specified in the WM_QUIT message.                     |
|                                                                              |
\*----------------------------------------------------------------------------*/
int PASCAL WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPSTR szCmdLine, int sw)
{
	(void)AllocConsole();
	AttachConsole(GetCurrentProcessId());
	freopen("CON", "w", stdout);
	freopen("CON", "w", stderr);
	freopen("CON", "r", stdin);
    MSG msg;

    /* Call initialization procedure */
    if(!AppInit(hInst,hPrev,sw))
        return FALSE;

    /*
    * Polling messages from event queue
    */
    for(;;)
    {
        while(PeekMessage(&msg, NULL, 0, 0,PM_REMOVE))
        {
            if(msg.message == WM_QUIT)
                break;  // Leave the PeekMessage while() loop

            if(TranslateAccelerator(ghwndApp, ghAccel, &msg))
                continue;

            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        if(msg.message == WM_QUIT)
            break;  // Leave the for() loop

        WaitMessage();
    }

    // Reached on WM_QUIT message
    CoUninitialize();
    return ((int) msg.wParam);
}


/*----------------------------------------------------------------------------*\
|   AppWndProc( hwnd, uiMessage, wParam, lParam )                              |
|                                                                              |
|   Description:                                                               |
|       The window proc for the app's main (tiled) window.  This processes all |
|       of the parent window's messages.                                       |
|                                                                              |
|   Arguments:                                                                 |
|       hwnd            window handle for the window                           |
|       msg             message number                                         |
|       wParam          message-dependent                                      |
|       lParam          message-dependent                                      |
|                                                                              |
|   Returns:                                                                   |
|       0 if processed, nonzero if ignored                                     |
|                                                                              |
\*----------------------------------------------------------------------------*/
static LONG WINAPI  AppWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    PAINTSTRUCT ps;
    HDC hdc;
    RECT rc;
    int cxBorder, cyBorder, cy;

    switch(msg)
    {
        case WM_CREATE:
            break;

        case WM_COMMAND:
            return AppCommand(hwnd,msg,wParam,lParam);

        case WM_INITMENU:
            // We can bring up a dialog if the graph is stopped
            EnableMenuItem((HMENU)wParam, MENU_DIALOG0, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG1, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG2, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG3, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG4, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG5, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG6, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG7, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG8, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOG9, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOGA, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOGB, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOGC, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOGD, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOGE, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_DIALOGF, MF_ENABLED);
            
            // do we want preview?
            CheckMenuItem((HMENU)wParam, MENU_PREVIEW, MF_CHECKED);
            // can toggle preview if not capturing
            EnableMenuItem((HMENU)wParam, MENU_PREVIEW, MF_GRAYED);

            EnableMenuItem((HMENU)wParam, MENU_VDEVICE0, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE1, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE2, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE3, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE4, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE5, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE6, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE7, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE8, MF_ENABLED);
            EnableMenuItem((HMENU)wParam, MENU_VDEVICE9, MF_ENABLED);

            break;


        case WM_INITMENUPOPUP:
            if(GetSubMenu(GetMenu(ghwndApp), 1) == (HMENU)wParam)
            {
                AddDevicesToMenu();
            }
            break;

        //
        // We're out of here!
        //
        case WM_DESTROY:

            IMonRelease(gcap.pmVideo);
            {
                for(int i = 0; i < NUMELMS(gcap.rgpmVideoMenu); i++)
                {
                    IMonRelease(gcap.rgpmVideoMenu[i]);
                }
            }

            PostQuitMessage(0);
            break;


        case WM_CLOSE:
            OnClose();
            break;

        case WM_ENDSESSION:
            if(wParam || (lParam & ENDSESSION_LOGOFF))
            {
                OnClose();
            }
            break;

        case WM_ERASEBKGND:
            break;

        // ESC will stop capture
        case WM_KEYDOWN:
            break;

        case WM_PAINT:
            hdc = BeginPaint(hwnd,&ps);

            // nothing to do
            EndPaint(hwnd,&ps);
            break;

        case WM_TIMER:
            // update our status bar with #captured, #dropped
            // if we've stopped capturing, don't do it anymore.  Some WM_TIMER
            // messages may come late, after we've destroyed the graph and
            // we'll get invalid numbers.

            break;

        case WM_SIZE:
            // make the preview window fit inside our window, taking up
            // all of our client area except for the status window at the
            // bottom
            GetClientRect(ghwndApp, &rc);
            cxBorder = GetSystemMetrics(SM_CXBORDER);
            cyBorder = GetSystemMetrics(SM_CYBORDER);
            cy = statusGetHeight() + cyBorder;
            MoveWindow(ghwndStatus, -cxBorder, rc.bottom - cy,
                rc.right + (2 * cxBorder), cy + cyBorder, TRUE);
            rc.bottom -= cy;
            // this is the video renderer window showing the preview
            //if(gcap.pVW)
            //    gcap.pVW->SetWindowPosition(0, 0, rc.right, rc.bottom);
            break;

        case WM_FGNOTIFY:
            // uh-oh, something went wrong while capturing - the filtergraph
            // will send us events like EC_COMPLETE, EC_USERABORT and the one
            // we care about, EC_ERRORABORT.
            if(gcap.pME)
            {
                LONG event;
				LONG_PTR l1, l2;
                HRESULT hrAbort = S_OK;
                BOOL bAbort = FALSE;
                while(gcap.pME->GetEvent(&event, &l1, &l2, 0) == S_OK)
                {
                    gcap.pME->FreeEventParams(event, l1, l2);
                    if(event == EC_ERRORABORT)
                    {
                        bAbort = TRUE;
                        hrAbort = static_cast<HRESULT>(l1);
                        continue;
                    }
                    else if(event == EC_DEVICE_LOST)
                    {
                            // Check if we have lost a capture filter being used.
                            // lParam2 of EC_DEVICE_LOST event == 1 indicates device added
                            //                                 == 0 indicates device removed
                            if(l2 == 0)
                            {
                                IBaseFilter *pf;
                                IUnknown *punk = (IUnknown *) l1;
                                if(S_OK == punk->QueryInterface(IID_IBaseFilter, (void **) &pf))
                                {
                                    if(AreComObjectsEqual(gcap.pVCap, pf))
                                    {
                                        pf->Release();
                                        bAbort = FALSE;
                                        TCHAR szError[100];
                                        HRESULT hr = StringCchCopy(szError, 100,
                                            TEXT("Stopping Capture (Device Lost). Select New Capture Device\0"));
                                        ErrMsg(szError);
                                        break;
                                    }
                                    pf->Release();
                                }
                            }
                    }
                } // end while
                if(bAbort)
                {
                        //if(gcap.fWantPreview)
                        {
                            BuildPreviewGraph();
                            StartPreview();
                        }
                        TCHAR szError[100];
                        HRESULT hr = StringCchPrintf(szError, 100, TEXT("ERROR during capture, error code=%08x\0"), hrAbort);
                        ErrMsg(szError);
                }
            }
            break;

        case WM_DEVICECHANGE:
            // We are interested in only device arrival & removal events
            if(DBT_DEVICEARRIVAL != wParam && DBT_DEVICEREMOVECOMPLETE != wParam)
                break;

            PDEV_BROADCAST_HDR pdbh = (PDEV_BROADCAST_HDR) lParam;
            if(pdbh->dbch_devicetype != DBT_DEVTYP_DEVICEINTERFACE)
            {
                break;
            }

            PDEV_BROADCAST_DEVICEINTERFACE pdbi = (PDEV_BROADCAST_DEVICEINTERFACE) lParam;
            // Check for capture devices.
            if(pdbi->dbcc_classguid != AM_KSCATEGORY_CAPTURE)
            {
                break;
            }

            // Check for device arrival/removal.
            if(DBT_DEVICEARRIVAL == wParam || DBT_DEVICEREMOVECOMPLETE == wParam)
            {
                gcap.fDeviceMenuPopulated = false;
            }
            break;

    }

    return (LONG) DefWindowProc(hwnd,msg,wParam,lParam);
}


// Make a graph builder object we can use for capture graph building
//
static BOOL MakeBuilder()
{
    // we have one already
    if(gcap.pBuilder)
        return TRUE;

    gcap.pBuilder = new ISampleCaptureGraphBuilder( );
    if( NULL == gcap.pBuilder )
    {
        return FALSE;
    }

    return TRUE;
}


// Make a graph object we can use for capture graph building
//
static BOOL MakeGraph()
{
    // we have one already
    if(gcap.pFg)
        return TRUE;

    HRESULT hr = CoCreateInstance(CLSID_FilterGraph, NULL, CLSCTX_INPROC,
                                  IID_IGraphBuilder, (LPVOID *)&gcap.pFg);

    return (hr == NOERROR) ? TRUE : FALSE;
}

// make sure the preview window inside our window is as big as the
// dimensions of captured video, or some capture cards won't show a preview.
// (Also, it helps people tell what size video they're capturing)
// We will resize our app's window big enough so that once the status bar
// is positioned at the bottom there will be enough room for the preview
// window to be w x h
//

static void ResizeWindow(int w, int h)
{
    static thread_local int gnRecurse = 0;

    RECT rcW, rcC;
    int xExtra, yExtra;
    int cyBorder = GetSystemMetrics(SM_CYBORDER);

    gnRecurse++;

    GetWindowRect(ghwndApp, &rcW);
    GetClientRect(ghwndApp, &rcC);
    xExtra = rcW.right - rcW.left - rcC.right;
    yExtra = rcW.bottom - rcW.top - rcC.bottom + cyBorder + statusGetHeight();

    rcC.right = w;
    rcC.bottom = h;
    SetWindowPos(ghwndApp, NULL, 0, 0, rcC.right + xExtra,
        rcC.bottom + yExtra, SWP_NOZORDER | SWP_NOMOVE);

    // we may need to recurse once.  But more than that means the window cannot
    // be made the size we want, trying will just stack fault.
    //
    if(gnRecurse == 1 && ((rcC.right + xExtra != rcW.right - rcW.left && w > GetSystemMetrics(SM_CXMIN)) ||
        (rcC.bottom + yExtra != rcW.bottom - rcW.top)))
        ResizeWindow(w,h);

    gnRecurse--;
}


// Tear down everything downstream of a given filter
static void RemoveDownstream(IBaseFilter *pf)
{
    IPin *pP=0, *pTo=0;
    ULONG u;
    IEnumPins *pins = NULL;
    PIN_INFO pininfo;

    if (!pf)
        return;

    HRESULT hr = pf->EnumPins(&pins);
    pins->Reset();

    while(hr == NOERROR)
    {
        hr = pins->Next(1, &pP, &u);
        if(hr == S_OK && pP)
        {
            pP->ConnectedTo(&pTo);
            if(pTo)
            {
                hr = pTo->QueryPinInfo(&pininfo);
                if(hr == NOERROR)
                {
                    if(pininfo.dir == PINDIR_INPUT)
                    {
                        RemoveDownstream(pininfo.pFilter);
                        gcap.pFg->Disconnect(pTo);
                        gcap.pFg->Disconnect(pP);
                        gcap.pFg->RemoveFilter(pininfo.pFilter);
                    }
                    pininfo.pFilter->Release();
                }
                pTo->Release();
            }
            pP->Release();
        }
    }

    if(pins)
        pins->Release();
}


// Tear down everything downstream of the capture filters, so we can build
// a different capture graph.  Notice that we never destroy the capture filters
// and WDM filters upstream of them, because then all the capture settings
// we've set would be lost.
//
static void TearDownGraph()
{
	SAFE_RELEASE(gcap.pGrabber);
    SAFE_RELEASE(gcap.pSink);
    SAFE_RELEASE(gcap.pRender);
    SAFE_RELEASE(gcap.pME);
    SAFE_RELEASE(gcap.pDF);
	SAFE_RELEASE(gcap.pVW);

    if(gcap.pVW && 0)
    {
        // stop drawing in our window, or we may get wierd repaint effects
        //gcap.pVW->put_Owner(NULL);
        //gcap.pVW->put_Visible(OAFALSE);
        gcap.pVW->Release();
        gcap.pVW = NULL;
    }

    // destroy the graph downstream of our capture filters
    if(gcap.pVCap)
        RemoveDownstream(gcap.pVCap);
    if(gcap.pVCap)
        gcap.pBuilder->ReleaseFilters();

    gcap.fCaptureGraphBuilt = FALSE;
    gcap.fPreviewGraphBuilt = FALSE;
    gcap.fPreviewFaked = FALSE;
}


// create the capture filters of the graph.  We need to keep them loaded from
// the beginning, so we can set parameters on them and have them remembered
//
static BOOL InitCapFilters()
{
    HRESULT hr=S_OK;
    BOOL f;

    f = MakeBuilder();
    if(!f)
    {
        ErrMsg(TEXT("Cannot instantiate graph builder"));
        return FALSE;
    }

    //
    // First, we need a Video Capture filter, and some interfaces
    //
    gcap.pVCap = NULL;

    if(gcap.pmVideo != 0)
    {
        IPropertyBag *pBag;
        gcap.wachFriendlyName[0] = 0;

        hr = gcap.pmVideo->BindToStorage(0, 0, IID_IPropertyBag, (void **)&pBag);
        if(SUCCEEDED(hr))
        {
            VARIANT var;
            var.vt = VT_BSTR;

            hr = pBag->Read(L"FriendlyName", &var, NULL);
            if(hr == NOERROR)
            {
                hr = StringCchCopyW(gcap.wachFriendlyName, sizeof(gcap.wachFriendlyName) / sizeof(gcap.wachFriendlyName[0]), var.bstrVal);
                SysFreeString(var.bstrVal);
            }

            pBag->Release();
        }

        hr = gcap.pmVideo->BindToObject(0, 0, IID_IBaseFilter, (void**)&gcap.pVCap);
    }

    if(gcap.pVCap == NULL)
    {
        ErrMsg(TEXT("Error %x: Cannot create video capture filter"), hr);
        goto InitCapFiltersFail;
    }

    //
    // make a filtergraph, give it to the graph builder and put the video
    // capture filter in the graph
    //

    f = MakeGraph();
    if(!f)
    {
        ErrMsg(TEXT("Cannot instantiate filtergraph"));
        goto InitCapFiltersFail;
    }

    hr = gcap.pBuilder->SetFiltergraph(gcap.pFg);
    if(hr != NOERROR)
    {
        ErrMsg(TEXT("Cannot give graph to builder"));
        goto InitCapFiltersFail;
    }

    // Add the video capture filter to the graph with its friendly name
    hr = gcap.pFg->AddFilter(gcap.pVCap, gcap.wachFriendlyName);
    if(hr != NOERROR)
    {
        ErrMsg(TEXT("Error %x: Cannot add vidcap to filtergraph"), hr);
        goto InitCapFiltersFail;
    }

    // Calling FindInterface below will result in building the upstream
    // section of the capture graph (any WDM TVTuners or Crossbars we might
    // need).

    // we use this interface to get the name of the driver
    // Don't worry if it doesn't work:  This interface may not be available
    // until the pin is connected, or it may not be available at all.
    // (eg: interface may not be available for some DV capture)
    hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                      &MEDIATYPE_Interleaved, gcap.pVCap,
                                      IID_IAMVideoCompression, (void **)&gcap.pVC);
    if(hr != S_OK)
    {
        hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                          &MEDIATYPE_Video, gcap.pVCap,
                                          IID_IAMVideoCompression, (void **)&gcap.pVC);
    }

    // !!! What if this interface isn't supported?
    // we use this interface to set the frame rate and get the capture size
    hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                      &MEDIATYPE_Interleaved,
                                      gcap.pVCap, IID_IAMStreamConfig, (void **)&gcap.pVSC);

    if(hr != NOERROR)
    {
        hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                          &MEDIATYPE_Video, gcap.pVCap,
                                          IID_IAMStreamConfig, (void **)&gcap.pVSC);
        if(hr != NOERROR)
        {
            // this means we can't set frame rate (non-DV only)
            ErrMsg(TEXT("Error %x: Cannot find VCapture:IAMStreamConfig"), hr);
        }
    }

    // potential debug output - what the graph looks like
    // DumpGraph(gcap.pFg, 1);

	AM_MEDIA_TYPE *pmt;

	// default capture format
	if (gcap.pVSC && gcap.pVSC->GetFormat(&pmt) == S_OK)
	{
		// DV capture does not use a VIDEOINFOHEADER
		if (pmt->formattype == FORMAT_VideoInfo)
		{
			// resize our window to the default capture size
			ResizeWindow(HEADER(pmt->pbFormat)->biWidth,
				ABS(HEADER(pmt->pbFormat)->biHeight));
		}
		DeleteMediaType(pmt);
	}

	return TRUE;

InitCapFiltersFail:
    FreeCapFilters();
    return FALSE;
}


// all done with the capture filters and the graph builder
//
static void FreeCapFilters()
{
    SAFE_RELEASE(gcap.pFg);
    if( gcap.pBuilder )
    {
        delete gcap.pBuilder;
        gcap.pBuilder = NULL;
    }
    SAFE_RELEASE(gcap.pVCap);
    SAFE_RELEASE(gcap.pVSC);
    SAFE_RELEASE(gcap.pVC);
    SAFE_RELEASE(gcap.pDlg);
}


// build the preview graph!
//
// !!! PLEASE NOTE !!!  Some new WDM devices have totally separate capture
// and preview settings.  An application that wishes to preview and then
// capture may have to set the preview pin format using IAMStreamConfig on the
// preview pin, and then again on the capture pin to capture with that format.
// In this sample app, there is a separate page to set the settings on the
// capture pin and one for the preview pin.  To avoid the user
// having to enter the same settings in 2 dialog boxes, an app can have its own
// UI for choosing a format (the possible formats can be enumerated using
// IAMStreamConfig) and then the app can programmatically call IAMStreamConfig
// to set the format on both pins.
//
static BOOL BuildPreviewGraph()
{
	HRESULT hr;

	// we have one already
	if (gcap.fPreviewGraphBuilt)
		return TRUE;

	// We don't have the necessary capture filters
	if (gcap.pVCap == NULL)
		return FALSE;

	// we already have another graph built... tear down the old one
	if (gcap.fCaptureGraphBuilt)
		TearDownGraph();

	//
	// Render the preview pin - even if there is not preview pin, the capture
	// graph builder will use a smart tee filter and provide a preview.
	//
	// !!! what about latency/buffer issues?

	// NOTE that we try to render the interleaved pin before the video pin, because
	// if BOTH exist, it's a DV filter and the only way to get the audio is to use
	// the interleaved pin.  Using the Video pin on a DV filter is only useful if
	// you don't want the audio.

	check(CoCreateInstance(CLSID_SampleGrabber, NULL, CLSCTX_INPROC_SERVER, IID_ISampleGrabber, (void**)&gcap.pGrabber));

	check(gcap.pGrabber->SetCallback(&gcap.callback, 1));
	check(gcap.pGrabber->SetBufferSamples(TRUE));

	check(gcap.pGrabber->QueryInterface(IID_IBaseFilter, (void**)&gcap.pVW));

	check(gcap.pFg->AddFilter(gcap.pVW, L"Sample Crapper"));

	check(gcap.pBuilder->RenderStream(NULL, NULL, gcap.pVCap, NULL, gcap.pVW));

    // All done.
    gcap.fPreviewGraphBuilt = TRUE;
    return TRUE;
fail:

	printf("hr %d\n", (int)hr);
	ErrMsg(TEXT("This graph cannot preview!"));
	gcap.fPreviewGraphBuilt = FALSE;
	return FALSE;
}


// Start previewing
//
static BOOL StartPreview()
{
	if (gcap.previewing)
		return TRUE;

	gcap.previewing = true;

    if(!gcap.fPreviewGraphBuilt)
        return FALSE;

    // run the graph
    IMediaControl *pMC = NULL;
    HRESULT hr = gcap.pFg->QueryInterface(IID_IMediaControl, (void **)&pMC);
    if(SUCCEEDED(hr))
    {
        hr = pMC->Run();
        if(FAILED(hr))
        {
            // stop parts that ran
            pMC->Stop();
        }
        pMC->Release();
    }
    if(FAILED(hr))
    {
        ErrMsg(TEXT("Error %x: Cannot run preview graph"), hr);
        return FALSE;
    }

    return TRUE;
}


// stop the preview graph
//
static BOOL StopPreview()
{
	if (!gcap.previewing)
		return TRUE;

	gcap.previewing = FALSE;

    // stop the graph
    IMediaControl *pMC = NULL;
    HRESULT hr = gcap.pFg->QueryInterface(IID_IMediaControl, (void **)&pMC);
    if(SUCCEEDED(hr))
    {
        hr = pMC->Stop();
        pMC->Release();
    }
    if(FAILED(hr))
    {
        ErrMsg(TEXT("Error %x: Cannot stop preview graph"), hr);
        return FALSE;
    }

    // get rid of menu garbage
    InvalidateRect(ghwndApp, NULL, TRUE);

    return TRUE;
}

// Let's talk about UI for a minute.  There are many programmatic interfaces
// you can use to program a capture filter or related filter to capture the
// way you want it to.... eg:  IAMStreamConfig, IAMVideoCompression,
// IAMCrossbar, IAMTVTuner, IAMTVAudio, IAMAnalogVideoDecoder, IAMCameraControl,
// IAMVideoProcAmp, etc.
//
// But you probably want some UI to let the user play with all these settings.
// For new WDM-style capture devices, we offer some default UI you can use.
// The code below shows how to bring up all of the dialog boxes supported
// by any capture filters.
//
// The following code shows you how you can bring up all of the
// dialogs supported by a particular object at once on a big page with lots
// of thumb tabs.  You do this by starting with an interface on the object that
// you want, and using ISpecifyPropertyPages to get the whole list, and
// OleCreatePropertyFrame to bring them all up.  This way you will get custom
// property pages a filter has, too, that are not one of the standard pages that
// you know about.  There are at least 9 objects that may have property pages.
// Your app already has 2 of the object pointers, the video capture filter and
// the audio capture filter (let's call them pVCap and pACap)
// 1.  The video capture filter - pVCap
// 2.  The video capture filter's capture pin - get this by calling
//     FindInterface(&PIN_CATEGORY_CAPTURE, pVCap, IID_IPin, &pX);
// 3.  The video capture filter's preview pin - get this by calling
//     FindInterface(&PIN_CATEGORY_PREVIEW, pVCap, IID_IPin, &pX);
// 4.  The audio capture filter - pACap
// 5.  The audio capture filter's capture pin - get this by calling
//     FindInterface(&PIN_CATEGORY_CAPTURE, pACap, IID_IPin, &pX);
// 6.  The crossbar connected to the video capture filter - get this by calling
//     FindInterface(NULL, pVCap, IID_IAMCrossbar, &pX);
// 7.  There is a possible second crossbar to control audio - get this by
//     looking upstream of the first crossbar like this:
//     FindInterface(&LOOK_UPSTREAM_ONLY, pX, IID_IAMCrossbar, &pX2);
// 8.  The TV Tuner connected to the video capture filter - get this by calling
//     FindInterface(NULL, pVCap, IID_IAMTVTuner, &pX);
// 9.  The TV Audio connected to the audio capture filter - get this by calling
//     FindInterface(NULL, pACap, IID_IAMTVAudio, &pX);
// 10. We have a helper class, CCrossbar, which makes the crossbar issue less
//     confusing.  In fact, although not supported here, there may be more than
//     two crossbars, arranged in many different ways.  An application may not
//     wish to have separate dialogs for each crossbar, but instead hide the
//     complexity and simply offer the user a list of inputs that can be chosen.
//     This list represents all the unique inputs from all the crossbars.
//     The crossbar helper class does this and offers that list as #10.  It is
//     expected that an application will either provide the crossbar dialogs
//     above (#6 and #7) OR provide the input list (this #10), but not both.
//     That would be confusing because if you select an input using dialog 6 or
//     7 the input list here in #10 won't know about your choice.
//
// Your last choice for UI is to make your own pages, and use the results of
// your custom page to call the interfaces programmatically.


static void MakeMenuOptions()
{
    HRESULT hr;
    HMENU hMenuSub = GetSubMenu(GetMenu(ghwndApp), 2); // Options menu

    // remove any old choices from the last device
	for (unsigned i = 0; i < 10 ; i++)
		RemoveMenu(hMenuSub, 0, MF_BYPOSITION);
    
    int zz = 0;
    gcap.iFormatDialogPos = -1;
    gcap.iSourceDialogPos = -1;
    gcap.iDisplayDialogPos = -1;
    gcap.iVCapDialogPos = -1;
    gcap.iVCapCapturePinDialogPos = -1;
    gcap.iVCapPreviewPinDialogPos = -1;

    // If this device supports the old legacy UI dialogs, offer them
    if(gcap.pDlg)
    {

        hr = gcap.pDlg->HasDialog(VfwCaptureDialog_Format);
        if(SUCCEEDED(hr) && S_FALSE != hr)
        {
            AppendMenu(hMenuSub, MF_STRING, MENU_DIALOG0 + zz, TEXT("Video Format..."));
            gcap.iFormatDialogPos = zz++;
        }
        hr = gcap.pDlg->HasDialog(VfwCaptureDialog_Source);
        if(SUCCEEDED(hr) && S_FALSE != hr)
        {
            AppendMenu(hMenuSub, MF_STRING, MENU_DIALOG0 + zz, TEXT("Video Source..."));
            gcap.iSourceDialogPos = zz++;
        }
        hr = gcap.pDlg->HasDialog(VfwCaptureDialog_Display);
        if(SUCCEEDED(hr) && S_FALSE != hr)
        {
            AppendMenu(hMenuSub, MF_STRING, MENU_DIALOG0 + zz, TEXT("Video Display..."));
            gcap.iDisplayDialogPos = zz++;
        }
    }

    // don't bother looking for new property pages if the old ones are supported
    // or if we don't have a capture filter
    if(gcap.pVCap == NULL || gcap.iFormatDialogPos != -1)
        return;

    // New WDM devices support new UI and new interfaces.
    // Your app can use some default property
    // pages for UI if you'd like (like we do here) or if you don't like our
    // dialog boxes, feel free to make your own and programmatically set
    // the capture options through interfaces like IAMCrossbar, IAMCameraControl
    // etc.

    // There are 9 objects that might support property pages.
    // Let's go through them.

    ISpecifyPropertyPages *pSpec;
    CAUUID cauuid;

    // 1. the video capture filter itself

    hr = gcap.pVCap->QueryInterface(IID_ISpecifyPropertyPages, (void **)&pSpec);
    if(SUCCEEDED(hr))
    {
        hr = pSpec->GetPages(&cauuid);
        if(SUCCEEDED(hr) && cauuid.cElems > 0)
        {
            AppendMenu(hMenuSub,MF_STRING,MENU_DIALOG0+zz, TEXT("Video Capture Filter..."));
            gcap.iVCapDialogPos = zz++;
            CoTaskMemFree(cauuid.pElems);
        }
        pSpec->Release();
    }

    // 2.  The video capture capture pin

    IAMStreamConfig *pSC;

    hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                      &MEDIATYPE_Interleaved,
                                      gcap.pVCap, IID_IAMStreamConfig, (void **)&pSC);
    if(FAILED(hr))
        hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                                      &MEDIATYPE_Video, gcap.pVCap,
                                      IID_IAMStreamConfig, (void **)&pSC);

    if(SUCCEEDED(hr))
    {
        hr = pSC->QueryInterface(IID_ISpecifyPropertyPages, (void **)&pSpec);
        if(SUCCEEDED(hr))
        {
            hr = pSpec->GetPages(&cauuid);
            if(SUCCEEDED(hr) && cauuid.cElems > 0)
            {
                AppendMenu(hMenuSub,MF_STRING,MENU_DIALOG0+zz, TEXT("Video Capture Pin..."));
                gcap.iVCapCapturePinDialogPos = zz++;
                CoTaskMemFree(cauuid.pElems);
            }
            pSpec->Release();
        }
        pSC->Release();
    }

    // 3.  The video capture preview pin.
    // This basically sets the format being previewed.  Typically, you
    // want to capture and preview using the SAME format, instead of having to
    // enter the same value in 2 dialog boxes.  For a discussion on this, see
    // the comment above the MakePreviewGraph function.

    hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_PREVIEW,
                                      &MEDIATYPE_Interleaved, gcap.pVCap,
                                      IID_IAMStreamConfig, (void **)&pSC);
    if(FAILED(hr))
        hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_PREVIEW,
                                          &MEDIATYPE_Video, gcap.pVCap,
                                          IID_IAMStreamConfig, (void **)&pSC);
    if(SUCCEEDED(hr))
    {
        hr = pSC->QueryInterface(IID_ISpecifyPropertyPages, (void **)&pSpec);
        if(SUCCEEDED(hr))
        {
            hr = pSpec->GetPages(&cauuid);
            if(SUCCEEDED(hr) && cauuid.cElems > 0)
            {
                AppendMenu(hMenuSub,MF_STRING,MENU_DIALOG0+zz,TEXT("Video Preview Pin..."));
                gcap.iVCapPreviewPinDialogPos = zz++;
                CoTaskMemFree(cauuid.pElems);
            }
            pSpec->Release();
        }
        pSC->Release();
    }
}

// Check the devices we're currently using and make filters for them
//
static void ChooseDevices(IMoniker *pmVideo)
{
#define VERSIZE 40
#define DESCSIZE 80

    int versize = VERSIZE;
    int descsize = DESCSIZE;
    WCHAR wachVer[VERSIZE]={0}, wachDesc[DESCSIZE]={0};
    TCHAR tachStatus[VERSIZE + DESCSIZE + 5]={0};


    // they chose a new device. rebuild the graphs
    if(gcap.pmVideo != pmVideo)
    {
        if(pmVideo)
        {
            pmVideo->AddRef();
        }

        IMonRelease(gcap.pmVideo);
        gcap.pmVideo = pmVideo;

        if(1)
            StopPreview();
        if(gcap.fCaptureGraphBuilt || gcap.fPreviewGraphBuilt)
            TearDownGraph();

        FreeCapFilters();
        InitCapFilters();

        if(1)   // were we previewing?
        {
            BuildPreviewGraph();
            StartPreview();
        }

        MakeMenuOptions();      // the UI choices change per device
    }

    // Set the check marks for the devices menu.
    int i;
    for(i = 0; i < NUMELMS(gcap.rgpmVideoMenu); i++)
    {
        if(gcap.rgpmVideoMenu[i] == NULL)
            break;

        CheckMenuItem(GetMenu(ghwndApp),
            MENU_VDEVICE0 + i,
            (S_OK == gcap.rgpmVideoMenu[i]->IsEqual(gcap.pmVideo)) ? MF_CHECKED : MF_UNCHECKED);
    }

    // Put the video driver name in the status bar - if the filter supports
    // IAMVideoCompression::GetInfo, that's the best way to get the name and
    // the version.  Otherwise use the name we got from device enumeration
    // as a fallback.
    if(gcap.pVC)
    {
        HRESULT hr = gcap.pVC->GetInfo(wachVer, &versize, wachDesc, &descsize,
                                       NULL, NULL, NULL, NULL);
        if(hr == S_OK)
        {
            // It's possible that the call succeeded without actually filling
            // in information for description and version.  If these strings
            // are empty, just display the device's friendly name.
            if(wcslen(wachDesc) && wcslen(wachVer))
            {
                hr = StringCchPrintf(tachStatus, VERSIZE + DESCSIZE + 5, TEXT("%s - %s\0"), wachDesc, wachVer);
                statusUpdateStatus(ghwndStatus, tachStatus);
                return;
            }
        }
    }

    // Since the GetInfo method failed (or the interface did not exist),
    // display the device's friendly name.
    statusUpdateStatus(ghwndStatus, gcap.wachFriendlyName);
}

static void ChooseDevices(TCHAR *szVideo)
{
	WCHAR wszVideo[1024];

    StringCchCopyN(wszVideo, NUMELMS(wszVideo), szVideo, NUMELMS(wszVideo)-1);
    wszVideo[1023] = 0;    // Null-terminate

    IBindCtx *lpBC=0;
	IMoniker *pmVideo = 0;

    HRESULT hr = CreateBindCtx(0, &lpBC);
    if(SUCCEEDED(hr))
    {
        DWORD dwEaten;
        hr = MkParseDisplayName(lpBC, wszVideo, &dwEaten, &pmVideo);

        lpBC->Release();
    }

    // Handle the case where the video capture device used for the previous session
    // is not available now.
    BOOL bFound = FALSE;

    if(pmVideo != NULL)
    {
        for(int i = 0; i < NUMELMS(gcap.rgpmVideoMenu); i++)
        {
            if(gcap.rgpmVideoMenu[i] != NULL &&
               S_OK == gcap.rgpmVideoMenu[i]->IsEqual(pmVideo))
            {
                bFound = TRUE;
                break;
            }
        }
    }

    if(!bFound)
    {
        if(gcap.iNumVCapDevices > 0)
        {
            IMonRelease(pmVideo);
            ASSERT(gcap.rgpmVideoMenu[0] != NULL);
            pmVideo = gcap.rgpmVideoMenu[0];
            pmVideo->AddRef();
        }
        else
            goto CleanUp;
    }

    ChooseDevices(pmVideo);

CleanUp:
    IMonRelease(pmVideo);
}


// put all installed video and audio devices in the menus
//
static void AddDevicesToMenu()
{
    if(gcap.fDeviceMenuPopulated)
    {
        return;
    }
    gcap.fDeviceMenuPopulated = true;
    gcap.iNumVCapDevices = 0;

    UINT    uIndex = 0;
    HMENU   hMenuSub;
    HRESULT hr;
    BOOL bCheck = FALSE;

    hMenuSub = GetSubMenu(GetMenu(ghwndApp), 1);        // Devices menu

    // Clean the sub menu
    int iMenuItems = GetMenuItemCount(hMenuSub);
    if(iMenuItems == -1)
    {
        ErrMsg(TEXT("Error Cleaning Devices Menu"));
        return;
    }
    else if(iMenuItems > 0)
    {
        for(int i = 0; i < iMenuItems; i++)
        {
            RemoveMenu(hMenuSub, 0, MF_BYPOSITION);
        }
    }

    for(int i = 0; i < NUMELMS(gcap.rgpmVideoMenu); i++)
    {
        IMonRelease(gcap.rgpmVideoMenu[i]);
    }

    // enumerate all video capture devices
    ICreateDevEnum *pCreateDevEnum=0;
    hr = CoCreateInstance(CLSID_SystemDeviceEnum, NULL, CLSCTX_INPROC_SERVER,
                          IID_ICreateDevEnum, (void**)&pCreateDevEnum);
    if(hr != NOERROR)
    {
        ErrMsg(TEXT("Error Creating Device Enumerator"));
        return;
    }

    IEnumMoniker *pEm=0;
    hr = pCreateDevEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory, &pEm, 0);
    if(hr != NOERROR)
    {
        ErrMsg(TEXT("Sorry, you have no video capture hardware.\r\n\r\n")
               TEXT("Video capture will not function properly."));
    }

    pEm->Reset();
    ULONG cFetched;
    IMoniker *pM;

    while(hr = pEm->Next(1, &pM, &cFetched), hr==S_OK)
    {
        IPropertyBag *pBag=0;

        hr = pM->BindToStorage(0, 0, IID_IPropertyBag, (void **)&pBag);
        if(SUCCEEDED(hr))
        {
            VARIANT var;
            var.vt = VT_BSTR;
            hr = pBag->Read(L"FriendlyName", &var, NULL);
            if(hr == NOERROR)
            {
                AppendMenu(hMenuSub, MF_STRING, MENU_VDEVICE0 + uIndex,
                    var.bstrVal);

                if(gcap.pmVideo != 0 && (S_OK == gcap.pmVideo->IsEqual(pM)))
                    bCheck = TRUE;

                CheckMenuItem(hMenuSub,  MENU_VDEVICE0 + uIndex,
                    (bCheck ? MF_CHECKED : MF_UNCHECKED));
                EnableMenuItem(hMenuSub, MENU_VDEVICE0 + uIndex, MF_ENABLED);
                bCheck = FALSE;

                SysFreeString(var.bstrVal);

                ASSERT(gcap.rgpmVideoMenu[uIndex] == 0);
                gcap.rgpmVideoMenu[uIndex] = pM;
                pM->AddRef();
            }
            pBag->Release();
        }

        pM->Release();
        uIndex++;
    }
    pEm->Release();

    gcap.iNumVCapDevices = uIndex;

    // separate the video and audio devices
    AppendMenu(hMenuSub, MF_SEPARATOR, 0, NULL);
}

/*----------------------------------------------------------------------------*\
|    AppCommand()
|
|    Process all of our WM_COMMAND messages.
\*----------------------------------------------------------------------------*/
static LONG PASCAL AppCommand(HWND hwnd, unsigned msg, WPARAM wParam, LPARAM lParam)
{
    HRESULT hr;
    int id = GET_WM_COMMAND_ID(wParam, lParam);

    switch(id)
    {
        // Our about box
        //
        case MENU_ABOUT:
            DialogBox(ghInstApp, MAKEINTRESOURCE(IDD_ABOUT), hwnd,
                (DLGPROC)AboutDlgProc);
            break;

        // Exit the application
        //
        case MENU_EXIT:
            PostMessage(hwnd,WM_CLOSE,0,0L);
            break;

        // pick which video capture device to use
        // pick which video capture device to use
        //
        case MENU_VDEVICE0:
        case MENU_VDEVICE1:
        case MENU_VDEVICE2:
        case MENU_VDEVICE3:
        case MENU_VDEVICE4:
        case MENU_VDEVICE5:
        case MENU_VDEVICE6:
        case MENU_VDEVICE7:
        case MENU_VDEVICE8:
        case MENU_VDEVICE9:
            ChooseDevices(gcap.rgpmVideoMenu[id - MENU_VDEVICE0]);
            break;

        // video format dialog
        //
        case MENU_DIALOG0:
        case MENU_DIALOG1:
        case MENU_DIALOG2:
        case MENU_DIALOG3:
        case MENU_DIALOG4:
        case MENU_DIALOG5:
        case MENU_DIALOG6:
        case MENU_DIALOG7:
        case MENU_DIALOG8:
        case MENU_DIALOG9:
        case MENU_DIALOGA:
        case MENU_DIALOGB:
        case MENU_DIALOGC:
        case MENU_DIALOGD:
        case MENU_DIALOGE:
        case MENU_DIALOGF:

            // they want the VfW format dialog
            if(id - MENU_DIALOG0 == gcap.iFormatDialogPos)
            {
                // this dialog will not work while previewing
                if(1)
                    StopPreview();
                HRESULT hrD;
                hrD = gcap.pDlg->ShowDialog(VfwCaptureDialog_Format, ghwndApp);

                // Sometimes bringing up the FORMAT dialog can result
                // in changing to a capture format that the current graph
                // can't handle.  It looks like that has happened and we'll
                // have to rebuild the graph.
                if(hrD == VFW_E_CANNOT_CONNECT)
                {
                    TearDownGraph();    // now we need to rebuild
                    // !!! This won't work if we've left a stranded h/w codec
                }

				// Resize our window to be the same size that we're capturing
				if (gcap.pVSC)
				{
					AM_MEDIA_TYPE *pmt;
					// get format being used NOW
					hr = gcap.pVSC->GetFormat(&pmt);

					// DV capture does not use a VIDEOINFOHEADER
					if (hr == NOERROR)
					{
						if (pmt->formattype == FORMAT_VideoInfo)
						{
							// resize our window to the new capture size
							ResizeWindow(HEADER(pmt->pbFormat)->biWidth,
								abs(HEADER(pmt->pbFormat)->biHeight));
						}
						DeleteMediaType(pmt);
					}
				}

                if(1)
                {
                    BuildPreviewGraph();
                    StartPreview();
                }
            }
            else if(id - MENU_DIALOG0 == gcap.iSourceDialogPos)
            {
                // this dialog will not work while previewing
                if(1)
                    StopPreview();

                gcap.pDlg->ShowDialog(VfwCaptureDialog_Source, ghwndApp);
                if(1)
                    StartPreview();
            }
            else if(id - MENU_DIALOG0 == gcap.iDisplayDialogPos)
            {
                // this dialog will not work while previewing
                if(1)
                    StopPreview();

                gcap.pDlg->ShowDialog(VfwCaptureDialog_Display, ghwndApp);
                if(1)
                    StartPreview();

                // now the code for the new dialogs
            }
            else if(id - MENU_DIALOG0 == gcap.iVCapDialogPos)
            {
                ISpecifyPropertyPages *pSpec;
                CAUUID cauuid;

                hr = gcap.pVCap->QueryInterface(IID_ISpecifyPropertyPages,
                    (void **)&pSpec);
                if(hr == S_OK)
                {
                    hr = pSpec->GetPages(&cauuid);

                    hr = OleCreatePropertyFrame(ghwndApp, 30, 30, NULL, 1,
                        (IUnknown **)&gcap.pVCap, cauuid.cElems,
                        (GUID *)cauuid.pElems, 0, 0, NULL);

                    CoTaskMemFree(cauuid.pElems);
                    pSpec->Release();
                }
            }
            else if(id - MENU_DIALOG0 == gcap.iVCapCapturePinDialogPos)
            {
                // You can change this pin's output format in these dialogs.
                // If the capture pin is already connected to somebody who's
                // fussy about the connection type, that may prevent using
                // this dialog(!) because the filter it's connected to might not
                // allow reconnecting to a new format. (EG: you switch from RGB
                // to some compressed type, and need to pull in a decoder)
                // I need to tear down the graph downstream of the
                // capture filter before bringing up these dialogs.
                // In any case, the graph must be STOPPED when calling them.
                if(1)
                    StopPreview();  // make sure graph is stopped

                // The capture pin that we are trying to set the format on is connected if
                // one of these variable is set to TRUE. The pin should be disconnected for
                // the dialog to work properly.
                if(gcap.fCaptureGraphBuilt || gcap.fPreviewGraphBuilt)
                {
                    TearDownGraph();    // graph could prevent dialog working
                }

                IAMStreamConfig *pSC;
                hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                    &MEDIATYPE_Interleaved, gcap.pVCap,
                    IID_IAMStreamConfig, (void **)&pSC);

                if(hr != NOERROR)
                    hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_CAPTURE,
                        &MEDIATYPE_Video, gcap.pVCap,
                        IID_IAMStreamConfig, (void **)&pSC);

                ISpecifyPropertyPages *pSpec;
                CAUUID cauuid;

                hr = pSC->QueryInterface(IID_ISpecifyPropertyPages,
                    (void **)&pSpec);

                if(hr == S_OK)
                {
                    hr = pSpec->GetPages(&cauuid);
                    hr = OleCreatePropertyFrame(ghwndApp, 30, 30, NULL, 1,
                        (IUnknown **)&pSC, cauuid.cElems,
                        (GUID *)cauuid.pElems, 0, 0, NULL);

                    // !!! What if changing output formats couldn't reconnect
                    // and the graph is broken?  Shouldn't be possible...

                    CoTaskMemFree(cauuid.pElems);
                    pSpec->Release();
                }

                pSC->Release();
                if(1)
                {
                    BuildPreviewGraph();
                    StartPreview();
                }
            }
            else if(id - MENU_DIALOG0 == gcap.iVCapPreviewPinDialogPos)
            {
                // this dialog may not work if the preview pin is connected
                // already, because the downstream filter may reject a format
                // change, so we better kill the graph. (EG: We switch from
                // capturing RGB to some compressed fmt, and need to pull in
                // a decompressor)
                if(1)
                {
                    StopPreview();
                    TearDownGraph();
                }

                IAMStreamConfig *pSC;

                // This dialog changes the preview format, so it might affect
                // the format being drawn.  Our app's window size is taken
                // from the size of the capture pin's video, not the preview
                // pin, so changing that here won't have any effect. All in all,
                // this probably won't be a terribly useful dialog in this app.
                hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_PREVIEW,
                                                  &MEDIATYPE_Interleaved, gcap.pVCap,
                                                  IID_IAMStreamConfig, (void **)&pSC);
                if (hr != NOERROR)
                {
                    hr = gcap.pBuilder->FindInterface(&PIN_CATEGORY_PREVIEW,
                        &MEDIATYPE_Video, gcap.pVCap,
                        IID_IAMStreamConfig, (void **)&pSC);
                }

                ISpecifyPropertyPages *pSpec;
                CAUUID cauuid;

                hr = pSC->QueryInterface(IID_ISpecifyPropertyPages,
                    (void **)&pSpec);
                if(hr == S_OK)
                {
                    hr = pSpec->GetPages(&cauuid);

                    hr = OleCreatePropertyFrame(ghwndApp, 30, 30, NULL, 1,
                        (IUnknown **)&pSC, cauuid.cElems,
                        (GUID *)cauuid.pElems, 0, 0, NULL);

                    CoTaskMemFree(cauuid.pElems);
                    pSpec->Release();
                }

                pSC->Release();

                if(1)
                {
                    BuildPreviewGraph();
                    StartPreview();
                }
            }
    }
    return 0L;
}


/*----------------------------------------------------------------------------*\
|   ErrMsg - Opens a Message box with a error message in it.  The user can     |
|            select the OK button to continue                                  |
\*----------------------------------------------------------------------------*/
static void ErrMsg(LPTSTR szFormat,...)
{
    static TCHAR szBuffer[2048]={0};
    const size_t NUMCHARS = sizeof(szBuffer) / sizeof(szBuffer[0]);
    const int LASTCHAR = NUMCHARS - 1;

    // Format the input string
    va_list pArgs;
    va_start(pArgs, szFormat);

    // Use a bounded buffer size to prevent buffer overruns.  Limit count to
    // character size minus one to allow for a NULL terminating character.
    HRESULT hr = StringCchVPrintf(szBuffer, NUMCHARS - 1, szFormat, pArgs);
    va_end(pArgs);

    // Ensure that the formatted string is NULL-terminated
    szBuffer[LASTCHAR] = TEXT('\0');

    MessageBox(ghwndApp, szBuffer, NULL,
               MB_OK | MB_ICONEXCLAMATION | MB_TASKMODAL);
}


/* AboutDlgProc()
 *
 * Dialog Procedure for the "about" dialog box.
 *
 */

static BOOL CALLBACK AboutDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch(msg)
    {
        case WM_COMMAND:
            EndDialog(hwnd, TRUE);
            return TRUE;

        case WM_INITDIALOG:
            return TRUE;
    }
    return FALSE;
}

static void OnClose()
{
    // Unregister device notifications
    if(ghDevNotify != NULL)
    {
        ASSERT(gpUnregisterDeviceNotification);
        gpUnregisterDeviceNotification(ghDevNotify);
        ghDevNotify = NULL;
    }

    // Destroy the filter graph and cleanup
    StopPreview();
    TearDownGraph();
    FreeCapFilters();
}