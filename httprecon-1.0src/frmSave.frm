VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save Fingerprints"
   ClientHeight    =   4092
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   3972
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4092
   ScaleWidth      =   3972
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1692
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   3735
      Begin VB.CheckBox chkUpload 
         Caption         =   "Submit fingerprint to project online database"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Note: Your fingerprints will be submitted to the project web site"
         Top             =   240
         Value           =   1  'Checked
         Width           =   3492
      End
      Begin VB.TextBox txtRemarks 
         Height          =   732
         Left            =   240
         MaxLength       =   127
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Example: ""Internal and behind a Squid proxy."""
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Optional remarks for fingerprint maintainer"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3372
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Cancel Database Update"
      Top             =   3600
      Width           =   1212
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Save Fingerprints"
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Height          =   1692
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtImplementationName 
         Height          =   285
         Left            =   240
         MaxLength       =   127
         TabIndex        =   1
         Text            =   "Apache 2.0.53"
         ToolTipText     =   "Example: Apache 2.0.53"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Apache 2.0.53"
         Height          =   252
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label Label5 
         Caption         =   "<name> <version> [details]"
         Height          =   252
         Left            =   1200
         TabIndex        =   11
         Top             =   480
         Width           =   2412
      End
      Begin VB.Label Label3 
         Caption         =   "Example:"
         Height          =   252
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Syntax:"
         Height          =   252
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Name of the httpd implementation you suspect."
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3492
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUpload_Click()
    If (chkUpload.Value = 0) Then
        txtRemarks.Enabled = False
        MsgBox "It is very sad that you do not want to participate to the project." & vbCrLf & _
            "Please submit new fingerprints, those will be added to the public" & vbCrLf & _
            "repository, to improve the quality of the enumeration.", vbInformation, "Help to improve"
    Else
        txtRemarks.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sImplementationName As String
    Dim sFullFingerprint As String
    Dim sRemarks As String
    
    sImplementationName = txtImplementationName.Text
    sRemarks = txtRemarks.Text

    Call SaveFingerprints(sImplementationName)
        
    If (Me.chkUpload.Value = 1) Then
        sFullFingerprint = GenerateFingerprintXML
        Call PostFingerprinToWebsite(sImplementationName, sRemarks, sFullFingerprint)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtImplementationName.Text = GetBanner(response_getexist)
    
    Me.txtImplementationName.SelStart = 0
    Me.txtImplementationName.SelLength = Len(txtImplementationName.Text)
    
    Me.txtRemarks.Text = "Target was " & frmMain.txtTargetHost.Text & ":" & frmMain.cboTargetPort.Text
End Sub

Private Sub txtImplementationName_GotFocus()
    txtImplementationName.SelStart = 0
    txtImplementationName.SelLength = Len(txtImplementationName.Text)
End Sub

Private Sub txtImplementationName_KeyUp(KeyCode As Integer, Shift As Integer)
    Call DisableButtons
End Sub

Private Sub DisableButtons()
    If (LenB(Trim$(txtImplementationName.Text)) = 0) Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtImplementationName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DisableButtons
End Sub
