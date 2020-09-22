Attribute VB_Name = "basCOMMON"
Option Explicit
Public Enum AppIcons
    iconCAMERA = 101
    iconDOCUMENT = 102
End Enum
Public Enum IMG_SIZE
    IMG_SIXTEEN = 16
    IMG_THIRTYTWO = 32
    IMG_ALREADYSET = 0
    IMG_CUSTOM = 1
End Enum
'------------------------------------------------------------
' Determines wheather or not a file already exists
' or not for the path/file name passed.
'------------------------------------------------------------
Function FileExists(filename As String) As Boolean
    On Error Resume Next
    Dim x As Long
    x = Len(Dir$(filename))
    If Err Or x = 0 Then FileExists = False Else FileExists = True
End Function
'------------------------------------------------------------
' Purpose:  Used to set a given form's Icon property
'                to an icon from the Resource File.
'                Note the use of AppIcons
' Parameters:
' Example:
' Date: July,21 1998 @ 19:25:18
'------------------------------------------------------------
Public Sub SetFormIcon(frm As Form, lngICON As AppIcons)
    On Error Resume Next
    frm.Icon = LoadResPicture(lngICON, vbResIcon)
End Sub
'------------------------------------------------------------
' Purpose:  Used to Add an image to a ImageList from the resource file.  Note.  AppIcons must be declared.
' Parameters:
' Example:
' Date: July,21 1998 @ 18:22:18
'------------------------------------------------------------
Public Sub AddImage(imgLIST As ImageList, resICONVAL As AppIcons, Optional imgSIZE As IMG_SIZE = IMG_ALREADYSET, Optional CustomHeight As Long = 16, Optional CustomWidth As Long = 16)
    On Error Resume Next
    With imgLIST
        If imgSIZE <> IMG_ALREADYSET Then
            If imgSIZE <> IMG_CUSTOM Then
                .ImageHeight = imgSIZE
                .ImageWidth = imgSIZE
            Else
                .ImageHeight = CustomHeight
                .ImageWidth = CustomWidth
            End If
        End If
        .ListImages.Add , , LoadResPicture(resICONVAL, vbResIcon)
    End With
End Sub
'------------------------------------------------------------
' Purpose:  Changes the size of icons within an ImageList at RunTime.
' Parameters:
' Example:
' Date: July,21 1998 @ 18:22:47
'------------------------------------------------------------
Public Sub ChangeImageSize(imgLIST As ImageList, imgSIZE As IMG_SIZE, Optional CustomHeight As Long = 16, Optional CustomWidth As Long = 16)
    On Error Resume Next
    With imgLIST
        If imgSIZE <> IMG_ALREADYSET Then
            If imgSIZE <> IMG_CUSTOM Then
                .ImageHeight = imgSIZE
                .ImageWidth = imgSIZE
            Else
                .ImageHeight = CustomHeight
                .ImageWidth = CustomHeight
            End If
        End If
    End With
End Sub
