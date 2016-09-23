//------------------------------------------------------------------------------
// File: AMCap.h
//
// Desc: DirectShow sample code - audio/video capture.
//
// Copyright (c) Microsoft Corporation.  All rights reserved.
//------------------------------------------------------------------------------


// Macros
#ifndef SAFE_RELEASE
#define SAFE_RELEASE(x) { if (x) x->Release(); x = NULL; }
#endif


extern "C"
{
    typedef BOOL (/* WINUSERAPI */ WINAPI *PUnregisterDeviceNotification)(
        IN HDEVNOTIFY Handle
        );

    typedef HDEVNOTIFY (/* WINUSERAPI */ WINAPI *PRegisterDeviceNotificationA)(
        IN HANDLE hRecipient,
        IN LPVOID NotificationFilter,
        IN DWORD Flags
        );

    typedef HDEVNOTIFY (/* WINUSERAPI */ WINAPI *PRegisterDeviceNotificationW)(
        IN HANDLE hRecipient,
        IN LPVOID NotificationFilter,
        IN DWORD Flags
        );
}

#define PRegisterDeviceNotification  PRegisterDeviceNotificationW


//
// Resource constants
//
#define ID_APP      1000

/* Menu Items */
#define MENU_EXIT           4
#define MENU_PREVIEW        15
#define MENU_VDEVICE0       16
#define MENU_VDEVICE1       17
#define MENU_VDEVICE2       18
#define MENU_VDEVICE3       19
#define MENU_VDEVICE4       20
#define MENU_VDEVICE5       21
#define MENU_VDEVICE6       22
#define MENU_VDEVICE7       23
#define MENU_VDEVICE8       24
#define MENU_VDEVICE9       25
#define MENU_ABOUT          36
#define MENU_DIALOG0        42
#define MENU_DIALOG1        43
#define MENU_DIALOG2        44
#define MENU_DIALOG3        45
#define MENU_DIALOG4        46
#define MENU_DIALOG5        47
#define MENU_DIALOG6        48
#define MENU_DIALOG7        49
#define MENU_DIALOG8        50
#define MENU_DIALOG9        51
#define MENU_DIALOGA        52
#define MENU_DIALOGB        53
#define MENU_DIALOGC        54
#define MENU_DIALOGD        55
#define MENU_DIALOGE        56
#define MENU_DIALOGF        57

// Dialogs
#define IDD_ABOUT               600

// defines for dialogs

// window messages
#define WM_FGNOTIFY WM_USER+1
