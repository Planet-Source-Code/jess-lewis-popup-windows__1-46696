VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BroswerToolSuite"
   ClientHeight    =   4575
   ClientLeft      =   3075
   ClientTop       =   1965
   ClientWidth     =   5820
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Login"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUserName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPassword"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtUserName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPassword"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdLogin"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdLogout"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "User Config"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxUsers"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "adoUserConfig"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdDeleteUser"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAddUser"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtAUserName"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtAPassword"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkChangeConfig"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Visual Control"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkNewWindows"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkPopups"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkBannerAds"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Misc. Control"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkMovies"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "chkActiveX"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "chkSound"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "chkCookies"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Block Sites"
      TabPicture(4)   =   "Form1.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkBlockSites"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "flxBlockSites"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "adoBlockedSites"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtBlockSites"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdAddBlockedSite"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdDeleteBlockedSite"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Block Words"
      TabPicture(5)   =   "Form1.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "adoBlockedWords"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "chkBlockWords"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdDeleteBlockedWord"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmdAddBlockedWord"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "txtBlockWords"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "flxBlockWords"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Limit Sites"
      TabPicture(6)   =   "Form1.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "flxLimitSites"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "adoLimitSites"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "chkLimitSites"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "cmdDeleteLimitSite"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cmdAddLimitSite"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "txtLimitSites"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Registration"
      TabPicture(7)   =   "Form1.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblRegistration"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "lblCustomerIdentifier"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "lblNotification"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "txtRegistration"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Text1"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).ControlCount=   5
      Begin VB.CheckBox chkChangeConfig 
         Caption         =   "Admin Privileges?"
         Height          =   375
         Left            =   -71520
         TabIndex        =   43
         Top             =   2355
         Width           =   1575
      End
      Begin VB.TextBox txtAPassword 
         Height          =   285
         Left            =   -73320
         TabIndex        =   42
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtAUserName 
         Height          =   285
         Left            =   -74880
         TabIndex        =   40
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73320
         TabIndex        =   39
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -72360
         TabIndex        =   38
         Top             =   2880
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBlockWords 
         Bindings        =   "Form1.frx":03EA
         Height          =   1215
         Left            =   -74880
         TabIndex        =   37
         Top             =   1200
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   1
         _Band(0)._MapCol(0)._Name=   "Word"
         _Band(0)._MapCol(0)._RSIndex=   0
      End
      Begin VB.TextBox txtLimitSites 
         Height          =   285
         Left            =   -74880
         TabIndex        =   35
         Top             =   2520
         Width           =   5055
      End
      Begin VB.CommandButton cmdAddLimitSite 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73320
         TabIndex        =   34
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteLimitSite 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -72360
         TabIndex        =   33
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtBlockWords 
         Height          =   285
         Left            =   -74880
         TabIndex        =   32
         Top             =   2520
         Width           =   5055
      End
      Begin VB.CommandButton cmdAddBlockedWord 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73320
         TabIndex        =   31
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteBlockedWord 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -72360
         TabIndex        =   30
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteBlockedSite 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -72360
         TabIndex        =   29
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdAddBlockedSite 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73320
         TabIndex        =   28
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtBlockSites 
         Height          =   285
         Left            =   -74880
         TabIndex        =   27
         Top             =   2520
         Width           =   5055
      End
      Begin MSAdodcLib.Adodc adoBlockedSites 
         Height          =   375
         Left            =   -71280
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         Height          =   375
         Left            =   1440
         TabIndex        =   25
         Top             =   2640
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBlockSites 
         Bindings        =   "Form1.frx":0408
         Height          =   1215
         Left            =   -74880
         TabIndex        =   24
         Top             =   1200
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   2
         _Band(0)._MapCol(0)._Name=   "User"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(0)._Hidden=   -1  'True
         _Band(0)._MapCol(1)._Name=   "Site"
         _Band(0)._MapCol(1)._RSIndex=   1
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "123456789"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtRegistration 
         Height          =   285
         Left            =   -73440
         TabIndex        =   20
         Text            =   "Enter your name here"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Text            =   "Enter your name here"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chkCookies 
         Caption         =   "Block Cookies"
         Height          =   495
         Left            =   -73920
         TabIndex        =   14
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CheckBox chkLimitSites 
         Caption         =   "Limit Navigation to only sites listed"
         Height          =   375
         Left            =   -73440
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkBlockWords 
         Caption         =   "Block Sites with listed words "
         Height          =   375
         Left            =   -73440
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkBlockSites 
         Caption         =   "Block Listed Sites"
         Height          =   375
         Left            =   -73440
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Block Background Sound and Music"
         Height          =   495
         Left            =   -73920
         TabIndex        =   9
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CheckBox chkActiveX 
         Caption         =   "Block ActiveX controls from installing"
         Height          =   495
         Left            =   -73920
         TabIndex        =   8
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CheckBox chkMovies 
         Caption         =   "Block Movies"
         Height          =   495
         Left            =   -73920
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkBannerAds 
         Caption         =   "Block Banner Ads"
         Height          =   495
         Left            =   -74040
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkPopups 
         Caption         =   "Block popups only"
         Height          =   495
         Left            =   -74040
         TabIndex        =   4
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CheckBox chkNewWindows 
         Caption         =   "Block all popups and new browser windows"
         Height          =   495
         Left            =   -74040
         TabIndex        =   3
         Top             =   960
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc adoBlockedWords 
         Height          =   375
         Left            =   -71280
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc adoLimitSites 
         Height          =   375
         Left            =   -71280
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLimitSites 
         Bindings        =   "Form1.frx":0426
         Height          =   1215
         Left            =   -74880
         TabIndex        =   36
         Top             =   1200
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   2
         _Band(0)._MapCol(0)._Name=   "User"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(0)._Hidden=   -1  'True
         _Band(0)._MapCol(1)._Name=   "Site"
         _Band(0)._MapCol(1)._RSIndex=   1
      End
      Begin MSAdodcLib.Adodc adoUserConfig 
         Height          =   375
         Left            =   -71040
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxUsers 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   41
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   2
         _Band(0)._MapCol(0)._Name=   "User"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(0)._Hidden=   -1  'True
         _Band(0)._MapCol(1)._Name=   "Site"
         _Band(0)._MapCol(1)._RSIndex=   1
      End
      Begin VB.Label lblNotification 
         Caption         =   "BrowserToolSuite has been used for 20 days.  There is only 1 day of free usage left.  Please register."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   23
         Top             =   2760
         Width           =   5055
      End
      Begin VB.Label lblCustomerIdentifier 
         Caption         =   "Unique Identifier"
         Height          =   375
         Left            =   -73440
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblRegistration 
         Caption         =   "Registration Number"
         Height          =   375
         Left            =   -73440
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblUserName 
         Caption         =   "User Name"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      Caption         =   "Current User:  Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdApply_Click()
    userPreferences(0) = chkNewWindows.Value
    userPreferences(1) = chkPopups.Value
    userPreferences(2) = chkBannerAds.Value
    userPreferences(3) = chkMovies.Value
    userPreferences(4) = chkActiveX.Value
    userPreferences(5) = chkSound.Value
    userPreferences(6) = chkBlockSites.Value
    userPreferences(7) = chkLimitSites.Value
    userPreferences(8) = chkBlockWords.Value
    userPreferences(9) = chkCookies.Value
    StorePreferences ("Default")
End Sub



Private Sub cmdAddBlockedSite_Click()
  If Len(txtBlockSites.Text) > 0 Then
    Call InsertRecord("tblBlockedSites")
    adoBlockedSites.Refresh
  Else
    MsgBox "Please enter a site name to add and then click Add.", vbInformation, App.ProductName
  End If
End Sub

Private Sub cmdAddBlockedWord_Click()
    If Len(txtBlockWords.Text) > 0 Then
    Call InsertRecord("tblBlockedWords")
    adoBlockedWords.Refresh
  Else
    MsgBox "Please enter a word to add and then click Add.", vbInformation, App.ProductName
  End If
End Sub

Private Sub cmdAddLimitSite_Click()
    If Len(txtLimitSites.Text) > 0 Then
    Call InsertRecord("tblLimitedSites")
    adoLimitSites.Refresh
  Else
    MsgBox "Please enter a site name to add and then click Add.", vbInformation, App.ProductName
  End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeleteBlockedSite_Click()
Dim doublecheck As Long
    If Len(flxBlockSites.Text) > 0 Then
        doublecheck = MsgBox("Are you sure that " & flxBlockSites & " should be deleted?", vbYesNo, App.ProductName)
        If doublecheck = vbYes Then
            Call DeleteRecord("tblBlockedSites")
            adoBlockedSites.Refresh
        End If
    Else
        MsgBox "Please select a site name from the list and then click Delete.", vbInformation, App.ProductName
    End If
End Sub

Private Sub cmdDeleteBlockedWord_Click()
Dim doublecheck As Long
    If Len(flxBlockWords.Text) > 0 Then
        doublecheck = MsgBox("Are you sure that " & flxBlockWords & " should be deleted?", vbYesNo, App.ProductName)
        If doublecheck = vbYes Then
            Call DeleteRecord("tblBlockedWords")
            adoBlockedWords.Refresh
        End If
    Else
        MsgBox "Please select a word from the list and then click Delete.", vbInformation, App.ProductName
    End If
End Sub

Private Sub cmdDeleteLimitSite_Click()
    Dim doublecheck As Long
    If Len(flxLimitSites.Text) > 0 Then
        doublecheck = MsgBox("Are you sure that " & flxLimitSites & " should be deleted?", vbYesNo, App.ProductName)
        If doublecheck = vbYes Then
            Call DeleteRecord("tblLimitedSites")
            adoLimitSites.Refresh
        End If
    Else
        MsgBox "Please select a site name from the list and then click Delete.", vbInformation, App.ProductName
    End If
End Sub

Private Sub cmdOK_Click()
    Call cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Form1.Caption = App.ProductName
    chkNewWindows.Value = userPreferences(0)
    chkPopups.Value = userPreferences(1)
    chkBannerAds.Value = userPreferences(2)
    chkMovies.Value = userPreferences(3)
    chkActiveX.Value = userPreferences(4)
    chkSound.Value = userPreferences(5)
    chkBlockSites.Value = userPreferences(6)
    chkLimitSites.Value = userPreferences(7)
    chkBlockWords.Value = userPreferences(8)
    chkCookies.Value = userPreferences(9)
    lblUser.Caption = "Current User:  " & currentUser
    flxBlockWords.ColHeaderCaption(0, 99) = "Word"
    Call SetDataSources
    flxBlockSites.ColWidth(0) = 5055
    flxBlockWords.ColWidth(0) = 5055
    flxLimitSites.ColWidth(0) = 5055
End Sub

Private Sub SetDataSources()
    adoBlockedSites.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BrowserToolSuite.mdb;Persist Security Info=False;Jet OLEDB:Database Password=XLtuIO9C"
    adoBlockedSites.RecordSource = "SELECT Site FROM tblBlockedSites WHERE theUser ='" & currentUser & "'"
    adoBlockedSites.Refresh
    
    adoBlockedWords.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BrowserToolSuite.mdb;Persist Security Info=False;Jet OLEDB:Database Password=XLtuIO9C"
    adoBlockedWords.RecordSource = "SELECT Word FROM tblBlockedWords WHERE theUser ='" & currentUser & "'"
    adoBlockedWords.Refresh
    
    adoLimitSites.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BrowserToolSuite.mdb;Persist Security Info=False;Jet OLEDB:Database Password=XLtuIO9C"
    adoLimitSites.RecordSource = "SELECT Site FROM tblLimitedSites WHERE theUser ='" & currentUser & "'"
    adoLimitSites.Refresh
    
End Sub

