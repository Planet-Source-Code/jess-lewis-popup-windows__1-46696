VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IObjectWithSite
Implements IOleCommandTarget

Private m_pUnkSite As IUnknownVB
Private WithEvents m_ie As InternetExplorer
Attribute m_ie.VB_VarHelpID = -1

Private Const IID_IWebBrowserApp = "{0002DF05-0000-0000-C000-000000000046}"
Private Const IID_IWebBrowser2 = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"


Private Sub IObjectWithSite_GetSite(ByVal priid As VBShellLib.REFIID, ppvObj As VBShellLib.VOID)
    m_pUnkSite.QueryInterface priid, ppvObj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As VBShellLib.IUnknownVB)
    
    On Error GoTo Errs
    
    If ObjPtr(pSite) = 0 Then
        CopyMemory m_ie, 0&, 4
        Exit Sub
    End If
    
    Set m_pUnkSite = pSite    'Save the site pointer for GetSite
    
    Dim pServiceProvider As IServiceProvider
    Set pServiceProvider = m_pUnkSite
    
    Dim clsidWebApp As GUID
    Dim clsidWebBrowser2 As GUID
    
    'Query service provider to get IWebBrowser2 (InternetExplorer)
    CLSIDFromString StrPtr(IID_IWebBrowserApp), clsidWebApp
    CLSIDFromString StrPtr(IID_IWebBrowser2), clsidWebBrowser2
    
    Set m_ie = pServiceProvider.QueryService(VarPtr(clsidWebApp), VarPtr(clsidWebBrowser2))
    
    Set pServiceProvider = Nothing
    If preferencesChecked = False Then Call CheckPreferences
    Exit Sub
Errs:
    Err.Raise Err.Number, "SetSite", Err.Description
End Sub



Private Sub IOleCommandTarget_Exec(ByVal pguidCmdGroup As VBShellLib.LPGUID, ByVal nCmdID As VBShellLib.DWORD, ByVal nCmdExecOpt As VBShellLib.DWORD, pvaIn As Variant, pvaOut As Variant)
Dim response As Variant

    frmMain.Show vbModal
End Sub

Private Sub IOleCommandTarget_QueryStatus(ByVal pguidCmdGroup As VBShellLib.LPGUID, ByVal cCmds As VBShellLib.ULONG, ByVal prgCmds As VBShellLib.LPOLECMD, ByVal pCmdText As VBShellLib.LPOLECMDTEXT)

    Dim i As Integer
       
    For i = 0 To cCmds - 1
        Dim cmd As OLECMD
            
        CopyMemory cmd, ByVal prgCmds + (Len(cmd) * i), Len(cmd)
        cmd.cmdf = OLECMDF_SUPPORTED Or OLECMDF_ENABLED
        CopyMemory ByVal prgCmds + (Len(cmd) * i), cmd, Len(cmd)
    
    Next i
    
End Sub



Private Sub m_ie_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
 If userPreferences(1) = 1 And Len(TargetFrameName) > 0 Then
        If pDisp.MenuBar = False Then
            Cancel = True
            m_ie.Quit
        End If
 End If
     
 
End Sub


Private Sub m_ie_NewWindow2(ppDisp As Object, Cancel As Boolean)
    If userPreferences(0) = 1 Then
        Cancel = True
    End If
    
  
End Sub


 




