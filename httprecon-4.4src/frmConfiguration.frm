VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConfiguration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmConfiguration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLongrequest 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboLongrequestLength 
         Height          =   315
         Left            =   120
         TabIndex        =   67
         Text            =   "1024"
         ToolTipText     =   "1024"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtLongrequestChar 
         Height          =   285
         Left            =   120
         MaxLength       =   1
         TabIndex        =   20
         Text            =   "a"
         ToolTipText     =   "a"
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image Image12 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":058A
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image11 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":0946
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label32 
         Caption         =   "The definition of the length of the long request in bytes which is used in the according test case. Suggested value: 1024"
         Height          =   495
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label31 
         Caption         =   $"frmConfiguration.frx":0D24
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label15 
         Caption         =   "req_longrequest_length"
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
         Left            =   480
         TabIndex        =   43
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "req_longrequest_char"
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
         Left            =   480
         TabIndex        =   42
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame fraResources 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboResourcesAttack 
         Height          =   315
         Left            =   120
         TabIndex        =   66
         Text            =   "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;"
         ToolTipText     =   "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;"
         Top             =   3240
         Width           =   6135
      End
      Begin VB.ComboBox cboResourcesNotavailable 
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Text            =   "/404test_.html"
         ToolTipText     =   "/404test_.html"
         Top             =   2040
         Width           =   6135
      End
      Begin VB.ComboBox cboResourcesAvailable 
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Text            =   "/"
         ToolTipText     =   "/"
         Top             =   840
         Width           =   6135
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":0DC9
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Image9 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":11DE
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":15E4
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label30 
         Caption         =   "The definition of the resource which shall be used within all requests fetching an existing object. Suggested value: /"
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label29 
         Caption         =   $"frmConfiguration.frx":19E9
         Height          =   495
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label Label28 
         Caption         =   $"frmConfiguration.frx":1A7A
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   2760
         Width           =   6135
      End
      Begin VB.Label Label12 
         Caption         =   "req_resource_attack"
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
         Left            =   480
         TabIndex        =   41
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "req_resource_notavailable"
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
         Left            =   480
         TabIndex        =   40
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "req_resource_available"
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
         Left            =   480
         TabIndex        =   39
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame fraProtocols 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   35
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboProtocolsWrong 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Text            =   "HTTP/9.8"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtProtocolsLegitimate 
         Height          =   285
         Left            =   120
         MaxLength       =   128
         TabIndex        =   18
         Text            =   "HTTP/1.1"
         ToolTipText     =   "HTTP/1.1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":1B1E
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":1F20
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label27 
         Caption         =   $"frmConfiguration.frx":231A
         Height          =   495
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label26 
         Caption         =   "The http protocol version which shall be used within the test case for wrong protocol definitions. Suggested value: HTTP/9.8"
         Height          =   495
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label9 
         Caption         =   "req_protocol_legitimate"
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
         Left            =   480
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "req_protocol_wrong"
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
         Left            =   480
         TabIndex        =   36
         Top             =   1560
         Width           =   2175
      End
   End
   Begin VB.Frame fraMethods 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboMethodsNotexisting 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   "TEST"
         ToolTipText     =   "TEST"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox cboMethodsNotallowed 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Text            =   "DELETE"
         ToolTipText     =   "DELETE"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Image Image16 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":23AB
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image15 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":27B5
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label25 
         Caption         =   $"frmConfiguration.frx":2BC9
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label24 
         Caption         =   $"frmConfiguration.frx":2C6E
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label7 
         Caption         =   "req_method_notexisting"
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
         Left            =   480
         TabIndex        =   34
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "req_method_notallowed"
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
         Left            =   480
         TabIndex        =   33
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraTests 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkTestAttack 
         Caption         =   "scan_test_attack*"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "This request may cause harm to the target service."
         Top             =   3120
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestWrongprotocol 
         Caption         =   "scan_test_wrongprotocol"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestNonexistingmethod 
         Caption         =   "scan_test_nonexistingmethod"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestWrontmethod 
         Caption         =   "scan_test_wrongmethod"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestOptions 
         Caption         =   "scan_test_options"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestHead 
         Caption         =   "scan_test_head"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestGetlong 
         Caption         =   "scan_test_getlong*"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "This request may cause harm to the target service."
         Top             =   960
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestGetnonexisting 
         Caption         =   "scan_test_getnonexisting"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.CheckBox chkTestGetexisting 
         Caption         =   "scan_test_getexisting"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   3000
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   4080
         Picture         =   "frmConfiguration.frx":2D12
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label23 
         Caption         =   "The test for getting an existing file is required."
         Height          =   495
         Left            =   4440
         TabIndex        =   53
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "As more test cases you activate, as higher the accuracy of the enumeration will be."
         Height          =   855
         Left            =   4440
         TabIndex        =   52
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "The test cases flagged with an * might cause harm to the target service and might be detected easily by security systems."
         Height          =   1215
         Left            =   4440
         TabIndex        =   51
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame fraTiming 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   6375
      Begin VB.TextBox txtTimingReceive 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "5000"
         ToolTipText     =   "5000"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtTimingSend 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "5000"
         ToolTipText     =   "5000"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtTimingConnect 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "5000"
         ToolTipText     =   "5000"
         Top             =   840
         Width           =   615
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":30C1
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":3494
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":3863
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label20 
         Caption         =   "Amount of time in milliseconds which shall be waited for a provoked response before aborting with a timeout. Suggested value: 5000"
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   6135
      End
      Begin VB.Label Label19 
         Caption         =   "Amount of time in milliseconds which shall be waited to send a full request before aborting with a timeout. Suggested value: 5000"
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label Label17 
         Caption         =   $"frmConfiguration.frx":3BFA
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label3 
         Caption         =   "req_timeout_receive"
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
         Left            =   480
         TabIndex        =   28
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "req_timeout_send"
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
         Left            =   480
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "req_timeout_connect"
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
         Left            =   480
         TabIndex        =   26
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame fraAgent 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cboAgentName 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Text            =   "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11"
         ToolTipText     =   "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11"
         Top             =   960
         Width           =   6135
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":3C84
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label33 
         Caption         =   "The definition of the agent name which shall be used within the http requests to announce the client."
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label18 
         Caption         =   "req_agent_name"
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
         Left            =   480
         TabIndex        =   45
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraStatistics 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   240
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtStatisticsHitpointsmin 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "7"
         ToolTipText     =   "7"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtStatisticsHitpointsmax 
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "14"
         ToolTipText     =   "14"
         Top             =   2280
         Width           =   615
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":4075
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   120
         Picture         =   "frmConfiguration.frx":447A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Amount of minimum required hitpoints per test case to reach 100 % in the matches. Suggested value: 7"
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label21 
         Caption         =   "Amount of maximum possible hitpoints per test case to set the level of 100 % in the matches. Suggested value: 14"
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   6135
      End
      Begin VB.Label Label6 
         Caption         =   "app_hitpoints_minimum"
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
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "app_hitpoints_maximum"
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
         Left            =   480
         TabIndex        =   30
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   3480
      Picture         =   "frmConfiguration.frx":4881
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancel Changes"
      Top             =   4440
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   615
      Left            =   2160
      Picture         =   "frmConfiguration.frx":4C19
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Save Settings"
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbsSettings 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Timing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tests"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Methods"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Protocols"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resources"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Longrequest"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Agent"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAgentName_Change()
    cboAgentName.Text = PreventEmptyInput(cboAgentName.Text, "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11")
End Sub

Private Sub cboLongrequestLength_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cboLongrequestLength_LostFocus()
    cboLongrequestLength.Text = AllowIntegersOnly(CLng(Val(cboLongrequestLength.Text)), 1, 65535, 1024)
End Sub

Private Sub cboMethodsNotallowed_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboMethodsNotallowed, KeyAscii, iLeftOff
End Sub

Private Sub cboMethodsNotallowed_LostFocus()
    cboMethodsNotallowed.Text = PreventEmptyInput(cboMethodsNotallowed.Text, "DELETE")
End Sub

Private Sub cboMethodsNotexisting_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboMethodsNotexisting, KeyAscii, iLeftOff
End Sub

Private Sub cboMethodsNotexisting_LostFocus()
    cboMethodsNotexisting.Text = PreventEmptyInput(cboMethodsNotexisting.Text, "TEST")
End Sub

Private Sub cboProtocolsWrong_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboProtocolsWrong, KeyAscii, iLeftOff
End Sub

Private Sub cboProtocolsWrong_LostFocus()
    cboProtocolsWrong.Text = PreventEmptyInput(cboProtocolsWrong.Text, "HTTP/9.8")
End Sub

Private Sub cboResourcesAttack_LostFocus()
    cboResourcesAttack.Text = PreventEmptyInput(cboResourcesAttack.Text, "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;")
End Sub

Private Sub cboResourcesAvailable_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboResourcesAvailable, KeyAscii, iLeftOff
End Sub

Private Sub cboResourcesAvailable_LostFocus()
    cboResourcesAvailable.Text = PreventEmptyInput(cboResourcesAvailable.Text, "/")
End Sub

Private Sub cboResourcesNotavailable_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboResourcesNotavailable, KeyAscii, iLeftOff
End Sub

Private Sub cboResourcesNotavailable_LostFocus()
    cboResourcesNotavailable.Text = PreventEmptyInput(cboResourcesNotavailable.Text, "/404test_.html")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call SaveConfiguration
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Configuration - " & Replace(app_configuration_filename, App.Path, vbNullString, , , 1)
    
    cboMethodsNotallowed.AddItem "DELETE"
    cboMethodsNotallowed.AddItem "PUT"
    cboMethodsNotallowed.AddItem "TRACE"
    cboMethodsNotallowed.AddItem "TRACK"
    cboMethodsNotallowed.AddItem "OPTIONS"
    cboMethodsNotallowed.AddItem "CONNECT"
    cboMethodsNotallowed.AddItem "PROPFIND"
    cboMethodsNotallowed.AddItem "PROPPATCH"
    cboMethodsNotallowed.AddItem "MKCOL"
    cboMethodsNotallowed.AddItem "COPY"
    cboMethodsNotallowed.AddItem "MOVE"
    cboMethodsNotallowed.AddItem "LOCK"
    cboMethodsNotallowed.AddItem "UNLOCK"
    
    cboMethodsNotexisting.AddItem "TEST"
    cboMethodsNotexisting.AddItem "FOO"
    cboMethodsNotexisting.AddItem "BLAH"
    cboMethodsNotexisting.AddItem "ABCDE"
    cboMethodsNotexisting.AddItem "QWERTY"
    cboMethodsNotexisting.AddItem Chr$(Rand(65, 90)) & Chr$(Rand(65, 90)) & Chr$(Rand(65, 90)) & Chr$(Rand(65, 90)) & Chr$(Rand(65, 90))
    
    cboProtocolsWrong.AddItem "HTTP/9.8"
    cboProtocolsWrong.AddItem "HTTP/1.9"
    cboProtocolsWrong.AddItem "HTTP/X.Y"
    
    cboResourcesAvailable.AddItem "/"
    cboResourcesAvailable.AddItem "/index.html"
    cboResourcesAvailable.AddItem "/index.htm"
    cboResourcesAvailable.AddItem "/index.php"
    cboResourcesAvailable.AddItem "/index.php4"
    cboResourcesAvailable.AddItem "/index.php5"
    cboResourcesAvailable.AddItem "/index.jsp"
    cboResourcesAvailable.AddItem "/default.html"
    cboResourcesAvailable.AddItem "/default.htm"
    cboResourcesAvailable.AddItem "/default.asp"
    cboResourcesAvailable.AddItem "/default.aspx"
    cboResourcesAvailable.AddItem "/default.jsp"
    
    cboResourcesNotavailable.AddItem "/404test_.html"
    cboResourcesNotavailable.AddItem "/foo_.html"
    cboResourcesNotavailable.AddItem "/blah_.html"
    cboResourcesNotavailable.AddItem "/abcde_.html"
    cboResourcesNotavailable.AddItem "/qwerty_.html"
    
    cboResourcesAttack.AddItem "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;"
    cboResourcesAttack.AddItem "../../etc/passwd"
    cboResourcesAttack.AddItem "/scripts/..%c1%1c../winnt/system32/cmd.exe?/c+dir"
    cboResourcesAttack.AddItem "/forum.php?user=<script>alert(document.cookie);</script>"
    cboResourcesAttack.AddItem "/forum.php?user=' OR 1;"
    
    cboLongrequestLength.AddItem "256"
    cboLongrequestLength.AddItem "512"
    cboLongrequestLength.AddItem "1024"
    cboLongrequestLength.AddItem "2048"
    cboLongrequestLength.AddItem "4096"
    cboLongrequestLength.AddItem "8192"
    cboLongrequestLength.AddItem "16384"
    
    cboAgentName.AddItem "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11"
    cboAgentName.AddItem "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506; InfoPath.2)"
    cboAgentName.AddItem "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 2.0.50727; InfoPath.1; .NET CLR 1.1.4322)"
    cboAgentName.AddItem "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1)"
    cboAgentName.AddItem "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; DT)"
    cboAgentName.AddItem "Opera/9.24 (Windows NT 5.1; U; en)"
    cboAgentName.AddItem "Mozilla/5.0 (Macintosh; U; PPC Mac OS X Mach-O; en-US; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11"
    
    Call FillConfiguration
End Sub

Private Sub tbsSettings_Click()
    Dim iIndex As Integer
    
    iIndex = tbsSettings.SelectedItem.Index

    fraTiming.Visible = False
    fraStatistics.Visible = False
    fraTests.Visible = False
    fraMethods.Visible = False
    fraProtocols.Visible = False
    fraResources.Visible = False
    fraLongrequest.Visible = False
    fraAgent.Visible = False

    If (iIndex = 1) Then
        fraTiming.Visible = True
    ElseIf (iIndex = 2) Then
        fraStatistics.Visible = True
    ElseIf (iIndex = 3) Then
        fraTests.Visible = True
    ElseIf (iIndex = 4) Then
        fraMethods.Visible = True
    ElseIf (iIndex = 5) Then
        fraProtocols.Visible = True
    ElseIf (iIndex = 6) Then
        fraResources.Visible = True
    ElseIf (iIndex = 7) Then
        fraLongrequest.Visible = True
    ElseIf (iIndex = 8) Then
        fraAgent.Visible = True
    End If
End Sub

Private Sub FillConfiguration()
    txtTimingConnect.Text = req_timeout_connect
    txtTimingSend.Text = req_timeout_send
    txtTimingReceive.Text = req_timeout_receive
    
    txtStatisticsHitpointsmin.Text = app_hitpoints_minimum
    txtStatisticsHitpointsmax.Text = app_hitpoints_maximum
    
    'Call SetCheckboxesTest(scan_test_getexisting, chkTestGetexisting)
    Call SetCheckboxesTest(scan_test_getnonexisting, chkTestGetnonexisting)
    Call SetCheckboxesTest(scan_test_getlong, chkTestGetlong)
    Call SetCheckboxesTest(scan_test_head, chkTestHead)
    Call SetCheckboxesTest(scan_test_options, chkTestOptions)
    Call SetCheckboxesTest(scan_test_wrongmethod, chkTestWrontmethod)
    Call SetCheckboxesTest(scan_test_nonexistingmethod, chkTestNonexistingmethod)
    Call SetCheckboxesTest(scan_test_wrongprotocol, chkTestWrongprotocol)
    Call SetCheckboxesTest(scan_test_attack, chkTestAttack)
    
    cboMethodsNotallowed.Text = req_method_notallowed
    cboMethodsNotexisting.Text = req_method_notexisting
    
    txtProtocolsLegitimate.Text = req_protocol_legitimate
    cboProtocolsWrong.Text = req_protocol_wrong
    
    cboResourcesAvailable.Text = req_resource_available
    cboResourcesNotavailable.Text = req_resource_notavailable
    cboResourcesAttack.Text = req_resource_attack
    
    cboLongrequestLength.Text = req_longrequest_length
    txtLongrequestChar.Text = req_longrequest_char
    
    cboAgentName.Text = req_agent_name
End Sub

Private Sub SetCheckboxesTest(ByRef iValue As Integer, ByRef cCheckbox As CheckBox)
    If (iValue = 0) Then
        cCheckbox.Value = 0
    Else
        cCheckbox.Value = 1
    End If
End Sub

Private Function GetCheckboxesTest(ByRef cCheckbox As CheckBox) As Integer
    If (cCheckbox.Value = 0) Then
        GetCheckboxesTest = 0
    Else
        GetCheckboxesTest = 1
    End If
End Function

Private Sub SaveConfiguration()
    req_timeout_connect = CInt(txtTimingConnect.Text)
    req_timeout_send = CInt(txtTimingSend.Text)
    req_timeout_receive = CInt(txtTimingReceive.Text)
    
    app_hitpoints_minimum = CInt(txtStatisticsHitpointsmin.Text)
    app_hitpoints_maximum = CInt(txtStatisticsHitpointsmax.Text)
    
    scan_test_getexisting = GetCheckboxesTest(chkTestGetexisting)
    scan_test_getnonexisting = GetCheckboxesTest(chkTestGetnonexisting)
    scan_test_getlong = GetCheckboxesTest(chkTestGetlong)
    scan_test_head = GetCheckboxesTest(chkTestHead)
    scan_test_options = GetCheckboxesTest(chkTestOptions)
    scan_test_wrongmethod = GetCheckboxesTest(chkTestWrontmethod)
    scan_test_nonexistingmethod = GetCheckboxesTest(chkTestNonexistingmethod)
    scan_test_wrongprotocol = GetCheckboxesTest(chkTestWrongprotocol)
    scan_test_attack = GetCheckboxesTest(chkTestAttack)
    
    req_method_notallowed = Trim(cboMethodsNotallowed.Text)
    req_method_notexisting = Trim(cboMethodsNotexisting.Text)
    
    req_protocol_legitimate = Trim(txtProtocolsLegitimate.Text)
    req_protocol_wrong = Trim(cboProtocolsWrong.Text)
    
    req_resource_available = cboResourcesAvailable.Text
    req_resource_notavailable = cboResourcesNotavailable.Text
    req_resource_attack = cboResourcesAttack.Text
    
    req_longrequest_length = CInt(cboLongrequestLength.Text)
    req_longrequest_char = txtLongrequestChar.Text
    
    req_agent_name = Trim(cboAgentName.Text)
    
    Call WriteConfigurationToFile(app_configuration_filename)
End Sub

Private Sub txtLongrequestChar_DblClick()
    txtLongrequestChar.Text = Chr$(Rand(97, 122))
End Sub

Private Sub txtLongrequestChar_LostFocus()
    txtLongrequestChar.Text = PreventEmptyInput(txtLongrequestChar.Text, "a")
End Sub

Private Function AllowIntegersOnly(ByRef lInput As Long, ByRef lMinimum As Long, ByRef lMaximum As Long, ByRef lDefault As Long)
    If LenB(lInput) = 0 Or lInput = 0 Then
        AllowIntegersOnly = lDefault
    Else
        If lInput < lMinimum Then
            AllowIntegersOnly = lMinimum
        ElseIf lInput > lMaximum Then
            AllowIntegersOnly = lMaximum
        Else
            AllowIntegersOnly = lInput
        End If
    End If
End Function

Private Function PreventEmptyInput(ByRef sInput As String, ByRef sDefault As String) As String
    If (LenB(sInput)) Then
        PreventEmptyInput = sInput
    Else
        PreventEmptyInput = sDefault
    End If
End Function

Private Sub txtProtocolsLegitimate_LostFocus()
    txtProtocolsLegitimate.Text = PreventEmptyInput(txtProtocolsLegitimate.Text, "HTTP/1.1")
End Sub

Private Sub txtStatisticsHitpointsmax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtStatisticsHitpointsmax_LostFocus()
    txtStatisticsHitpointsmax.Text = AllowIntegersOnly(CLng(Val(txtStatisticsHitpointsmax.Text)), 1, 99, 14)
End Sub

Private Sub txtStatisticsHitpointsmin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtStatisticsHitpointsmin_LostFocus()
    txtStatisticsHitpointsmin.Text = AllowIntegersOnly(CLng(Val(txtStatisticsHitpointsmin.Text)), 1, 99, 7)
End Sub

Private Sub txtTimingConnect_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTimingConnect_LostFocus()
    txtTimingConnect.Text = AllowIntegersOnly(CLng(Val(txtTimingConnect.Text)), 50, 30000, 5000)
End Sub

Private Sub txtTimingReceive_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTimingReceive_LostFocus()
    txtTimingReceive.Text = AllowIntegersOnly(CLng(Val(txtTimingReceive.Text)), 50, 30000, 5000)
End Sub

Private Sub txtTimingSend_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTimingSend_LostFocus()
    txtTimingSend.Text = AllowIntegersOnly(CLng(Val(txtTimingSend.Text)), 50, 30000, 5000)
End Sub
