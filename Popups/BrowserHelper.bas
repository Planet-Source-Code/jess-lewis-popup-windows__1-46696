Attribute VB_Name = "BrowserHelper"
Option Explicit

Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pClsid As GUID) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function StringFromCLSID Lib "ole32.dll" (pClsid As GUID, lpszProgID As Long) As Long
Public BlockJustPopupWindows As Boolean
Public BlockAllNewWindows As Boolean
Public userPreferences(2) As Integer
Public preferencesChecked As Boolean

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Const MAX_PATH = 260
Public Const MF_STRING = &H0&
Public Const MF_ENABLED = &H0&
Public Const MF_SEPARATOR = &H800&


Public Type OLEMENUGROUPWIDTHS
    widths(0 To 5) As Long
End Type

Public Sub CheckPreferences()

    userPreferences(0) = GetSetting("PopupWindows", "NewWindows", "Available", Default:=0)
    userPreferences(1) = GetSetting("PopupWindows", "Popups", "Available", Default:=0)
    
    preferencesChecked = True
End Sub




