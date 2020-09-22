VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Cam Portal"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
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
   ScaleHeight     =   4665
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMIN 
      Caption         =   "&Minimize"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDELETE 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   3855
   End
   Begin VB.TextBox txtNAME 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
   End
   Begin MSComctlLib.ImageList imgSMALL 
      Left            =   4200
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvLIST 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblLABEL 
      Caption         =   "URL:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblLABEL 
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdADD_Click()
    On Error Resume Next
    Dim itm As ListItem
    If Me.txtNAME.Text <> "" And Me.txtURL.Text <> "" Then
        Set itm = lvLIST.ListItems.Add(, , Me.txtNAME.Text, , 1)
        itm.SubItems(1) = Me.txtURL.Text
        SaveList
    End If
End Sub
Private Sub cmdDELETE_Click()
    On Error Resume Next
    Dim itm As ListItem
    Set itm = lvLIST.SelectedItem
    If Not itm Is Nothing Then
        lvLIST.ListItems.Remove itm.Index
        SaveList
    End If
End Sub
Private Sub cmdMIN_Click()
    On Error Resume Next
    Me.WindowState = vbMinimized
End Sub
Private Sub Form_Load()
    On Error Resume Next
    SetFormIcon Me, iconCAMERA
    AddImage imgSMALL, iconCAMERA, IMG_SIXTEEN
    InitlvLIST
    FilllvLIST
End Sub
Private Sub InitlvLIST()
    On Error GoTo ErrorInitlvLIST
    With lvLIST
        .View = lvwReport
        .HideSelection = False
        .FullRowSelect = True
        Set .SmallIcons = imgSMALL
        .ColumnHeaders.Add , , "NAME", 1600
        .ColumnHeaders.Add , , "URL", 3100
    End With
    Exit Sub
ErrorInitlvLIST:
    MsgBox Err & ":Error in call to InitlvLIST()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub FilllvLIST()
    On Error GoTo ErrorFilllvLIST
    Dim fn As Long, fNAME As String, l As String
    Dim cNAME As String, cURL As String, itm As ListItem
    fNAME = App.Path & "\cam.dat"
    If FileExists(fNAME) Then
        fn = FreeFile
        Open fNAME For Input As #fn
        Do While Not EOF(fn)
            Line Input #fn, l
            cNAME = GetToken(l, "^")
            cURL = l
            Set itm = lvLIST.ListItems.Add(, , cNAME, , 1)
            itm.SubItems(1) = cURL
        Loop
        Close #fn
    End If
    Exit Sub
ErrorFilllvLIST:
    If fn <> 0 Then Close #fn
    MsgBox Err & ":Error in call to FilllvLIST()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub SaveList()
    On Error GoTo ErrorSaveList
    Dim itm As ListItem, f As Long
    Dim fNAME As String
    f = FreeFile
    fNAME = App.Path & "\cam.dat"
    Open fNAME For Output As #f
    For Each itm In lvLIST.ListItems
        Print #f, itm.Text & "^" & itm.SubItems(1)
    Next
    Close #f
    Exit Sub
ErrorSaveList:
    MsgBox Err & ":Error in call to SaveList()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim frm As Form
    SaveList
    For Each frm In Forms
        If frm.Name <> Me.Name Then Unload frm
    Next
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Dim frm As Form
    If Me.WindowState = vbMinimized Then
        Me.Caption = "Document - Untitled"
        SetFormIcon Me, iconDOCUMENT
        For Each frm In Forms
            If frm.Name <> Me.Name Then frm.Visible = False
        Next
    Else
        Me.Caption = "Web Cam Portal"
        SetFormIcon Me, iconCAMERA
        For Each frm In Forms
            If frm.Name <> Me.Name Then frm.Visible = True
        Next
    End If
End Sub
Private Sub lvLIST_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error Resume Next
    If NewString <> "" Then
        SaveList
    Else
        Cancel = True
    End If
End Sub
Private Sub lvLIST_DblClick()
    On Error Resume Next
    Dim itm As ListItem, frm As frmCAM
    Set itm = lvLIST.SelectedItem
    If Not itm Is Nothing Then
        Set frm = New frmCAM
        frm.LoadIt itm.Text, itm.SubItems(1)
        frm.Show
    End If
End Sub
