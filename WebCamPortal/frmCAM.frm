VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCAM 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CAM"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   1920
   End
   Begin MSComctlLib.ProgressBar prgBAR 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin WebCamPortal.Downloader dwnLOAD 
      Left            =   0
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picCAM 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   120
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   4830
   End
   Begin VB.Label lblWAIT 
      Caption         =   "Waiting 10 sec(s)..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frmCAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mURL As String
Private mNAME As String
Public Sub LoadIt(cNAME As String, cURL As String)
    On Error GoTo ErrorLoadIt
    Me.Caption = cNAME
    mURL = cURL
    mNAME = cNAME
    Me.dwnLOAD.BeginDownload mURL, App.Path & "\" & mNAME & ".jpg"
    Exit Sub
ErrorLoadIt:
    MsgBox Err & ":Error in call to LoadIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub dwnLOAD_DownloadComplete(MaxBytes As Long, SaveFile As String)
    On Error Resume Next
    If MaxBytes <> 0 Then
        Set Me.picCAM.Picture = LoadPicture(SaveFile)
        Me.prgBAR.Visible = False
        Me.lblWAIT.Visible = True
        Me.prgBAR.Value = Me.prgBAR.Min
        Me.Timer.Enabled = True
    End If
End Sub
Private Sub dwnLOAD_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    If MaxBytes <> 0 Then
        With Me.prgBAR
            .Max = MaxBytes
            .Value = CurBytes
        End With
    End If
End Sub
Private Sub Form_Load()
    On Error Resume Next
    SetFormIcon Me, iconCAMERA
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If FileExists(App.Path & "\" & mNAME & ".jpg") Then Kill App.Path & "\" & mNAME & ".jpg"
End Sub
Private Sub Timer_Timer()
    On Error Resume Next
    Static i As Long
    If i < 10 Then
        i = i + 1
        Me.lblWAIT.Caption = "Waiting " & 10 - i & " sec(s)..."
    Else
        Me.Timer.Enabled = False
        Me.dwnLOAD.BeginDownload mURL, App.Path & "\" & mNAME & ".jpg"
        Me.prgBAR.Visible = True
        Me.lblWAIT.Visible = False
        Me.lblWAIT.Caption = "Waiting 10 sec(s)..."
        i = 0
    End If
End Sub
