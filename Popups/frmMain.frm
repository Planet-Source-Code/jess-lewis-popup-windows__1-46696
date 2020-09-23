VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Popup Windows"
   ClientHeight    =   4590
   ClientLeft      =   1860
   ClientTop       =   1890
   ClientWidth     =   5070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   540
         Left            =   150
         Picture         =   "frmMain.frx":030A
         ScaleHeight     =   337.12
         ScaleMode       =   0  'User
         ScaleWidth      =   337.12
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
      Begin VB.Line LineDark 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   830
         X2              =   4680
         Y1              =   2030
         Y2              =   2030
      End
      Begin VB.Label Label6 
         Caption         =   $"frmMain.frx":0614
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   840
         TabIndex        =   8
         Top             =   2160
         Width           =   3870
      End
      Begin VB.Label Label5 
         Caption         =   "Popup Windows Version 1.0"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   3885
      End
      Begin VB.Label Label4 
         Caption         =   $"frmMain.frx":06AE
         ForeColor       =   &H00000000&
         Height          =   1170
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   3885
      End
      Begin VB.Line LineWhite 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   850
         X2              =   4680
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin VB.CheckBox chkNewWindows 
      Caption         =   "Block all new browser windows"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CheckBox chkPopups 
      Caption         =   "Block popups only"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdApply_Click()
    userPreferences(0) = chkNewWindows.Value
    userPreferences(1) = chkPopups.Value
    Call StorePreferences
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOK_Click()
    Call cmdApply_Click
    Unload Me
End Sub



Private Sub Form_Load()
    Call RetrievePreferences
    chkNewWindows.Value = userPreferences(0)
    chkPopups.Value = userPreferences(1)
End Sub

Private Sub RetrievePreferences()

    userPreferences(0) = GetSetting("PopupWindows", "NewWindows", "Available", Default:=0)
    userPreferences(1) = GetSetting("PopupWindows", "Popups", "Available", Default:=0)
      
End Sub

Private Sub StorePreferences()


    SaveSetting "PopupWindows", "NewWindows", "Available", userPreferences(0)
    SaveSetting "PopupWindows", "Popups", "Available", userPreferences(1)
   
End Sub


