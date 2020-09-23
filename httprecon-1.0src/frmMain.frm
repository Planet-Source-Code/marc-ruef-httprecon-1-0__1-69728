VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "httprecon"
   ClientHeight    =   6960
   ClientLeft      =   132
   ClientTop       =   708
   ClientWidth     =   8052
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   8052
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgScanSaveAs 
      Left            =   4920
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
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
      DefaultExt      =   "HTML Report (*.html)|*.html"
      DialogTitle     =   "Save Report As"
      FileName        =   "report.html"
      Filter          =   "HTML Report (*.html)|*.html"
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   4440
      Top             =   0
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      DefaultExt      =   "Fingerprint Scan Files (*.fps)|*.fps"
      DialogTitle     =   "Open Fingerprint Scan Files"
      Filter          =   "Fingerprint Scan Files (*.fps)|*.fps|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.ImageList imlHttpdIcons 
      Left            =   7320
      Top             =   6240
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   80
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
      EndProperty
   End
   Begin VB.TextBox txtResponses 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
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
            Caption         =   "GET exist"
            Object.ToolTipText     =   "GET / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET long request"
            Object.ToolTipText     =   "GET /aaa(...) HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET non-existent"
            Object.ToolTipText     =   "GET /404test.html HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "GET wrong protocol"
            Object.ToolTipText     =   "GET / HTTP/9.8"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HEAD exist"
            Object.ToolTipText     =   "HEAD / HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OPTIONS common"
            Object.ToolTipText     =   "OPTIONS * HTTP/1.1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "DELETE exist"
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
   Begin VB.Frame fraResults 
      Caption         =   "Results"
      Height          =   2652
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   7812
      Begin MSComctlLib.ListView lsvResults 
         Height          =   2292
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7572
         _ExtentX        =   13356
         _ExtentY        =   4043
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
   End
   Begin VB.Frame fraTarget 
      Caption         =   "Target"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7812
      Begin VB.ComboBox cboTargetPort 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Text            =   "80"
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
         Left            =   600
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
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "http://"
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   492
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewItem 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpemItem 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAsItem 
         Caption         =   "&Save As..."
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
   Begin VB.Menu mnuFingerprinting 
      Caption         =   "&Fingerprinting"
      Begin VB.Menu mnuFingerprintingAnalyzeItem 
         Caption         =   "&Analyze"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFingerprintingSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFingerprintingOnlineDBItem 
         Caption         =   "Online Fingerprint Database"
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
         Caption         =   "&Generate HTML Report"
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
      Begin VB.Menu mnuHelpHomepageItem 
         Caption         =   "httprecon &home page"
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

Private Sub cboTargetPort_Change()
    Dim sInput As String
    
    sInput = cboTargetPort.Text
    
    If (LenB(sInput) = 0) Then
        cboTargetPort.Text = 80
    ElseIf (sInput > 65535) Then
        cboTargetPort.Text = 65535
    End If
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
    
    Call InitializeDirectories
    Call InitializeFiles

    frmMain.cboTargetPort.AddItem ("80")
    frmMain.cboTargetPort.AddItem ("81")
    frmMain.cboTargetPort.AddItem ("82")
    frmMain.cboTargetPort.AddItem ("800")
    frmMain.cboTargetPort.AddItem ("8000")
    frmMain.cboTargetPort.AddItem ("8080")
    frmMain.cboTargetPort.AddItem ("8081")
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        If Me.WindowState <> vbMinimized Then
            If Me.Height < 6000 Then
                Me.Height = 6000
            End If
            
            If Me.Width < 7000 Then
                Me.Width = 7000
            End If
        End If
    
        Me.fraTarget.Width = frmMain.Width - 360
        Me.cmdAnalyze.Left = Me.fraTarget.Width - Me.cmdAnalyze.Width - 120
        
        Me.tbsViews.Width = Me.fraTarget.Width
        Me.txtResponses.Width = Me.fraTarget.Width - 240
        Me.tbsViews.Height = (frmMain.Height - Me.fraTarget.Height) / 2 - 480
        Me.txtResponses.Height = Me.tbsViews.Height - 480
        
        Me.fraResults.Top = Me.tbsViews.Top + Me.tbsViews.Height + 120
        Me.fraResults.Width = Me.fraTarget.Width
        Me.lsvResults.Width = Me.txtResponses.Width
        Me.fraResults.Height = Me.tbsViews.Height - 360
        Me.lsvResults.Height = Me.fraResults.Height - 360
    End If
End Sub

Private Sub lsvResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ListViewSort(lsvResults, ColumnHeader, (lsvResults.SortOrder + 1) Mod 2)
End Sub

Private Sub mnuFileExitItem_Click()
    Unload Me
End Sub

Private Sub mnuFileNewItem_Click()
    Call ResetAll
End Sub

Private Sub mnuFileOpemItem_Click()
    Dim sFileName As String
    Dim sFileContent As String
    
    cdgOpen.InitDir = App.Path
    cdgOpen.ShowOpen
    If (cdgOpen.CancelError = False) Then
        sFileName = cdgOpen.FileName
            
        If LenB(sFileName) Then
            If (Dir$(sFileName, 16) <> "") Then
                Call ResetAll
                
                sFileContent = ReadFile(sFileName)
                Call ReadFingerprintXML(sFileContent)
                Call AnalyzeFingerprintsAndShowResult
                Me.Caption = APP_NAME & " - " & Mid$(sFileName, InStrRev(sFileName, "\", , vbBinaryCompare) + 1)
            End If
        End If
    End If
End Sub

Private Sub mnuFileSaveAsItem_Click()
    Dim sFileName As String
    Dim sOutput As String
    Dim lMetaCodeCount As Long
    Dim lCount As Long
    Dim sOverride As String
    
    cdgScanSaveAs.InitDir = App.Path
    
    On Error Resume Next
    cdgScanSaveAs.ShowSave
    sFileName = cdgScanSaveAs.FileName
    If (LenB(sFileName)) Then
        If (Dir$(sFileName, 16) <> "") Then
            sOverride = MsgBox(sFileName & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Scan Save As")
        Else
            sOverride = 6
        End If
        
        If (sOverride = 6) Then
            Open sFileName For Output As #1
                Print #1, GenerateFingerprintXML()
            Close
        End If
    End If
End Sub

Private Sub mnuFingerprintingAnalyzeItem_Click()
    Call ServerAnalysis
End Sub

Private Sub mnuFingerprintingOnlineDBItem_Click()
    Call ShellExecute(frmMain.hwnd, "Open", PROJECT_WEBDB, "", App.Path, 1)
End Sub

Private Sub mnuFingerprintingSaveFingerprintItem_Click()
    frmSave.Show vbModal, frmMain
End Sub

Private Sub mnuHelpAboutItem_Click()
    frmAbout.Show vbModal, frmMain
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
    
    On Error Resume Next
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
