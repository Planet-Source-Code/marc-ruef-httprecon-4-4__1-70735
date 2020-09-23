VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "httprecon"
   ClientHeight    =   7260
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   8052
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   8052
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgOpenScanlist 
      Left            =   3960
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Scanlist Files (*.scl)|*.scl"
      DialogTitle     =   "Open Scanlist Files"
      Filter          =   "Scanlist Files (*.scl)|*.scl|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgOpenScan 
      Left            =   4440
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Fingerprint Scan Files (*.fps)|*.fps"
      DialogTitle     =   "Open Fingerprint Scan Files"
      Filter          =   "Fingerprint Scan Files (*.fps)|*.fps|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   6960
      Width           =   8055
      _ExtentX        =   14203
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Ready."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgSaveAsScan 
      Left            =   4920
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Fingerprint Scan Files (*.fps)|*.fps"
      DialogTitle     =   "Save Fingerprint Scan Files"
      FileName        =   "127-0-0-1.fps"
      Filter          =   "Fingerprints Scan Files (*.fps)|*.fps|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgReportSaveAs 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "HTML Report (*.html)|*.html"
      DialogTitle     =   "Save Report As"
      FileName        =   "report.html"
      Filter          =   "HTML Report (*.html)|*.html"
   End
   Begin MSComctlLib.ImageList imlHttpdIcons 
      Left            =   7440
      Top             =   4320
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   101
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E43
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1143
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":137D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1777
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2130
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2217
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":243E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2931
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F49
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":325A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":361E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3703
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B07
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5305
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B07
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6655
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":81CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8623
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A81
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9227
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":95E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A594
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B829
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BBDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BF7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C398
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C791
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CBCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D346
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D73E
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF48
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E352
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E776
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB59
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EF7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F3AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F786
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB45
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF45
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1034D
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":106EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E51
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":111F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":119F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1220F
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1263D
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":129C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1347A
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1389D
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14064
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":144A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1487A
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15026
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15484
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1584E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtResponses 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2535
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1440
      Width           =   7575
   End
   Begin MSComctlLib.TabStrip tbsViews 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13780
      _ExtentY        =   5313
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET existing"
            Object.ToolTipText     =   "GET / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET long request"
            Object.ToolTipText     =   "GET /aaa(...) HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET non-existing"
            Object.ToolTipText     =   "GET /404test_.html HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET wrong protocol"
            Object.ToolTipText     =   "GET / HTTP/9.8"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HEAD existing"
            Object.ToolTipText     =   "HEAD / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OPTIONS common"
            Object.ToolTipText     =   "OPTIONS / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "DELETE existing"
            Object.ToolTipText     =   "DELETE / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TEST method"
            Object.ToolTipText     =   "TEST / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Attack Request"
            Object.ToolTipText     =   "GET <attack_request> HTTP/1.1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7812
      Begin VB.ComboBox cboScheme 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Usually: http (non-encrypted)"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboTargetPort 
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Text            =   "80"
         ToolTipText     =   "Usually: 80 (http)"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "&Analyze"
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         ToolTipText     =   "Analyze Web Server"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTargetHost 
         Height          =   285
         Left            =   1200
         MaxLength       =   255
         TabIndex        =   1
         Text            =   "127.0.0.1"
         ToolTipText     =   "Example: www.computec.ch"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   ":"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComctlLib.ListView lsvResults 
      CausesValidation=   0   'False
      Height          =   2175
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   7575
      _ExtentX        =   13356
      _ExtentY        =   3831
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "imlHttpdIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "STRING"
         Object.Width           =   512
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "STRING"
         Text            =   "Name"
         Object.Width           =   7410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "NUMBER"
         Text            =   "Hits"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "NUMBER"
         Text            =   "Match %"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.TextBox txtFingerprint 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   7575
   End
   Begin MSComctlLib.TabStrip tbsResults 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   7815
      _ExtentX        =   13780
      _ExtentY        =   4678
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Full Matchlist"
            Object.ToolTipText     =   "Full List of Matches"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Full Fingerprint Details"
            Object.ToolTipText     =   "Full Fingerprint Details"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenScanlistItem 
         Caption         =   "Open Scanlist..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenScanItem 
         Caption         =   "&Open Scan..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAsScanItem 
         Caption         =   "&Save As Scan..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuConfiguration 
      Caption         =   "&Configuration"
      Begin VB.Menu mnuConfigurationEditItem 
         Caption         =   "&Edit Settings..."
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuConfigurationSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigurationReloadItem 
         Caption         =   "&Reload Configuration"
         Shortcut        =   +{F5}
      End
   End
   Begin VB.Menu mnuFingerprinting 
      Caption         =   "Finger&printing"
      Begin VB.Menu mnuFingerprintingAnalyzeItem 
         Caption         =   "&Analyze (network access)"
      End
      Begin VB.Menu mnuFingerprintingReanalyzeItem 
         Caption         =   "&Re-Analyze (without network)"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFingerprintingSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFingerprintingOnlineDBItem 
         Caption         =   "Online Fingerprint Database..."
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFingerprintingSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFingerprintingSaveFingerprintItem 
         Caption         =   "&Save Fingerprint..."
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnuReporting 
      Caption         =   "&Reporting"
      Begin VB.Menu mnuReportingGenerateReportItem 
         Caption         =   "&Generate HTML Report..."
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAboutItem 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpCheckForUpdatesItem 
         Caption         =   "Check for Updates..."
      End
      Begin VB.Menu mnuHelpHomepageItem 
         Caption         =   "httprecon &Home Page..."
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboScheme_Click()
    If (cboScheme.Text = "http://") Then
        Call ChangeSSLMode(False)
    Else
        Call ChangeSSLMode(True)
    End If
End Sub

Private Sub cboScheme_KeyUp(KeyCode As Integer, Shift As Integer)
    If (cboScheme.Text = "http://") Then
        Call ChangeSSLMode(False)
    Else
        Call ChangeSSLMode(True)
    End If
End Sub

Private Sub cboTargetPort_Change()
    Dim sInput As String
    
    sInput = cboTargetPort.Text
    
    Call ChangeSSLMode(False)
    If (LenB(sInput) = 0) Then
        cboTargetPort.Text = 80
    ElseIf (sInput > 65535) Then
        cboTargetPort.Text = 65535
    ElseIf (sInput = 443) Then
        Call ChangeSSLMode(True)
    ElseIf (sInput = 8443) Then
        Call ChangeSSLMode(True)
    End If
End Sub

Private Sub cboTargetPort_Click()
    Call cboTargetPort_Change
End Sub

Private Sub cboTargetPort_GotFocus()
    cboTargetPort.SelStart = 0
    cboTargetPort.SelLength = Len(cboTargetPort.Text)
End Sub

Private Sub cboTargetPort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select

    Static iLeftOff As Long
    ComboAutoComplete cboTargetPort, KeyAscii, iLeftOff
End Sub

Private Sub cboTargetPort_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
        Call ServerAnalysis
    End If
End Sub

Private Sub cmdAnalyze_Click()
    Call ServerAnalysis
End Sub

Private Sub Form_Load()
    frmMain.Caption = APP_NAME
    
    cboScheme.AddItem ("http://")
    cboScheme.AddItem ("https://")
    
    cboTargetPort.AddItem ("80")
    cboTargetPort.AddItem ("81")
    cboTargetPort.AddItem ("82")
    cboTargetPort.AddItem ("443")
    cboTargetPort.AddItem ("800")
    cboTargetPort.AddItem ("888")
    cboTargetPort.AddItem ("2301")
    cboTargetPort.AddItem ("8000")
    cboTargetPort.AddItem ("8001")
    cboTargetPort.AddItem ("8080")
    cboTargetPort.AddItem ("8081")
    cboTargetPort.AddItem ("8443")
    cboTargetPort.AddItem ("8888")
    
    Randomize
    Call LoadConfigFromFile
    
    txtTargetHost.Text = scan_targethost
    cboTargetPort.Text = scan_targetport
    If (scan_targetsecure = 1) Then
        Call ChangeSSLMode(True)
    Else
        Call ChangeSSLMode(False)
    End If
    
    Call InitializeDirectories
    Call InitializeFiles

    Call ChangeStatusBarReady
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        If WindowState <> vbMinimized Then
            If Height < 6000 Then
                Height = 6000
            End If
            
            If Width < 7000 Then
                Width = 7000
            End If
        End If
    
        fraTarget.Width = frmMain.Width - 360
        cmdAnalyze.Left = fraTarget.Width - cmdAnalyze.Width - 120
        
        tbsViews.Width = fraTarget.Width
        txtResponses.Width = fraTarget.Width - 240
        tbsViews.Height = (frmMain.Height - fraTarget.Height - stbStatus.Height) / 2 - 480
        txtResponses.Height = tbsViews.Height - 480
        
        tbsResults.Top = tbsViews.Top + tbsViews.Height + 120
        tbsResults.Width = fraTarget.Width
        
        lsvResults.Width = txtResponses.Width
        lsvResults.Top = tbsResults.Top + 360
        txtFingerprint.Width = lsvResults.Width
        txtFingerprint.Top = lsvResults.Top
        
        tbsResults.Height = tbsViews.Height - 360
        lsvResults.Height = tbsResults.Height - 480
        txtFingerprint.Height = lsvResults.Height
    End If
End Sub

Private Sub lsvResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewSort(lsvResults, ColumnHeader, (lsvResults.SortOrder + 1) Mod 2)
End Sub

Private Sub mnuConfigurationEditItem_Click()
    frmConfiguration.Show vbModal, frmMain
End Sub

Private Sub mnuConfigurationReloadItem_Click()
    Call LoadConfigFromFile(app_configuration_filename)
End Sub

Private Sub mnuFileExitItem_Click()
    Unload Me
End Sub

Private Sub mnuFileNewItem_Click()
    Call ResetAll
End Sub

Private Sub mnuFileOpenScanItem_Click()
    Dim sFileName As String
    Dim sFileContent As String
    
    cdgOpenScan.InitDir = App.Path
    
    On Error GoTo Cancel
    cdgOpenScan.ShowOpen
    sFileName = cdgOpenScan.FileName
    
    If LenB(sFileName) Then
        If (Dir$(sFileName, 16) <> "") Then
            Call ResetAll
            
            sFileContent = ReadFile(sFileName)
            Call ReadFingerprintXML(sFileContent)
            Call AnalyzeFingerprintsAndShowResult
            
            frmMain.Caption = APP_NAME & " - " & scan_targethost & ":" & scan_targetport & " (" & Mid$(sFileName, InStrRev(sFileName, "\", , vbBinaryCompare) + 1) & ")"
            frmMain.txtTargetHost = scan_targethost
            frmMain.cboTargetPort = scan_targetport
            If (scan_targetsecure) Then
                Call ChangeSSLMode(True)
            Else
                Call ChangeSSLMode(False)
            End If
        End If
    End If

Cancel:
End Sub

Private Sub mnuFileOpenScanlistItem_Click()
    Dim sFileName As String
    Dim sFileContent As String
    Dim sScanListItems() As String
    Dim iScanListItemsCount As Integer
    Dim i As Integer
    Dim sReportPath As String
    Dim sReportFileName As String
    
    cdgOpenScanlist.InitDir = App.Path
    
    On Error GoTo Cancel
    cdgOpenScanlist.ShowOpen
    sFileName = cdgOpenScanlist.FileName
    
    If LenB(sFileName) Then
        If (Dir$(sFileName, 16) <> "") Then
            sFileContent = ReadFile(sFileName)
            sScanListItems = Split(sFileContent, vbCrLf, , vbBinaryCompare)
            iScanListItemsCount = UBound(sScanListItems)
            
            sReportPath = BrowseForFolder(Me, "Choose the destination directory for report files (html export and scan fingerprint).")
            
            If (LenB(sReportPath)) Then
                For i = 0 To iScanListItemsCount
                    If (LenB(sScanListItems(i))) Then
                        Call ResetAll
                        
                        If (Left$(sScanListItems(i), 8) = "https://") Then
                            scan_targetsecure = 1
                            Call ChangeSSLMode(True)
                        Else
                            scan_targetsecure = 0
                            Call ChangeSSLMode(False)
                        End If
                        
                        scan_targetport = ExtractTargetPort(sScanListItems(i))
                        frmMain.cboTargetPort = scan_targetport
                        
                        scan_targethost = SanitizeHostInput(sScanListItems(i))
                        frmMain.txtTargetHost = scan_targethost
                        
                        Call ServerAnalysis
                        
                        'This is for training mode only
                        'frmSave.Show vbModal, frmMain
                        
                        sReportFileName = sReportPath & "\" & StringToFileName(scan_targethost & ":" & scan_targetport) & ".html"
                        On Error Resume Next
                        Open sReportFileName For Output As #1
                            Print #1, GenerateHtmlReport()
                        Close
                        
                        sReportFileName = sReportPath & "\" & StringToFileName(scan_targethost & ":" & scan_targetport) & ".fps"
                        Open sReportFileName For Output As #1
                            Print #1, GenerateFingerprintXML(True)
                        Close
                        
                        Call ChangeStatusBar("Scanlist with " & iScanListItemsCount & " items finished. Ready.")
                    End If
                Next i
            End If
        End If
    End If

Cancel:
End Sub

Private Sub mnuFileSaveAsScanItem_Click()
    Dim sFileName As String
    Dim sOutput As String
    Dim lMetaCodeCount As Long
    Dim lCount As Long
    Dim sOverride As String
    
    cdgSaveAsScan.InitDir = App.Path
    cdgSaveAsScan.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".fps"
    
    On Error GoTo Cancel
    cdgSaveAsScan.ShowSave
    sFileName = cdgSaveAsScan.FileName
    
    If (LenB(sFileName)) Then
        If (Dir$(sFileName, 16) <> "") Then
            sOverride = MsgBox(sFileName & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Scan Save As")
        Else
            sOverride = 6
        End If
        
        If (sOverride = 6) Then
            Open sFileName For Output As #1
                Print #1, GenerateFingerprintXML(True)
            Close
            
            frmMain.Caption = APP_NAME & " - " & scan_targethost & ":" & scan_targetport & " (" & Mid$(sFileName, InStrRev(sFileName, "\", , vbBinaryCompare) + 1) & ")"
        End If
    End If

Cancel:
End Sub

Private Sub mnuFingerprintingAnalyzeItem_Click()
    Call ServerAnalysis
End Sub

Private Sub mnuFingerprintingOnlineDBItem_Click()
    Call ShellExecute(frmMain.hwnd, "Open", PROJECT_WEBDB, "", App.Path, 1)
End Sub

Private Sub mnuFingerprintingReanalyzeItem_Click()
    Call DisableElements
    Call AnalyzeFingerprintsAndShowResult
    Call EnableElements
End Sub

Private Sub mnuFingerprintingSaveFingerprintItem_Click()
    frmSave.Show vbModal, frmMain
End Sub

Private Sub mnuHelpAboutItem_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuHelpCheckForUpdatesItem_Click()
    Call OpenUpdateWebsite
End Sub

Private Sub mnuHelpHomepageItem_Click()
    Call OpenProjectWebsite
End Sub

Private Sub mnuReportingGenerateReportItem_Click()
    Dim sFileName As String
    Dim sOutput As String
    Dim lMetaCodeCount As Long
    Dim lCount As Long
    Dim sOverride As String
    
    cdgReportSaveAs.InitDir = App.Path
    cdgReportSaveAs.FileName = StringToFileName(scan_targethost & ":" & scan_targetport) & ".html"
    
    On Error GoTo Cancel
    cdgReportSaveAs.ShowSave
    sFileName = cdgReportSaveAs.FileName
    
    If (LenB(sFileName)) Then
        If (Dir$(sFileName, 16) <> "") Then
            sOverride = MsgBox(sFileName & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Report Save As")
        Else
            sOverride = 6
        End If
        
        If (sOverride = 6) Then
            Open sFileName For Output As #1
                Print #1, GenerateHtmlReport()
            Close
        End If
    
        Call ShellExecute(frmMain.hwnd, "Open", sFileName, "", App.Path, 1)
    End If

Cancel:
End Sub

Private Sub tbsResults_Click()
    Dim iIndex As Integer
    
    iIndex = frmMain.tbsResults.SelectedItem.Index

    If (iIndex = 1) Then
        frmMain.lsvResults.Visible = True
        frmMain.txtFingerprint.Visible = False
    ElseIf (iIndex = 2) Then
        frmMain.txtFingerprint.Visible = True
        frmMain.lsvResults.Visible = False
    End If
End Sub

Private Sub tbsViews_Click()
    Call FillResponses
End Sub

Private Sub txtTargetHost_GotFocus()
    txtTargetHost.SelStart = 0
    txtTargetHost.SelLength = Len(txtTargetHost.Text)
End Sub

Private Sub txtTargetHost_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) Then
        Call ServerAnalysis
    End If
End Sub

Private Sub txtTargetHost_LostFocus()
    Dim sNewTarget As String
    
    sNewTarget = txtTargetHost.Text
    sNewTarget = SanitizeHostInput(sNewTarget)
    txtTargetHost.Text = sNewTarget
End Sub
