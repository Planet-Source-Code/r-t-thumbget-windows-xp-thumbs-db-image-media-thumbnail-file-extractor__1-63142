VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThumbGet"
   ClientHeight    =   1095
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1095
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblDrag 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":000C
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sFilePath As String
Private sFileName As String
Private sFileData As String
Private sThumbData() As String

Private Sub DoProcess()
Dim i As Integer
    Me.Caption = "Loading"
    sFileData = GetFileContents(sFilePath & "\" & sFileName)
    
    Me.Caption = "Splitting"
    sThumbData = Split(sFileData, "ÿØÿà")
    
    Me.Caption = "Saving"
    For i = 1 To UBound(sThumbData)
        SetFileContents sFilePath & "\thumbs\", i & ".jpg", "ÿØÿà" & sThumbData(i)
    Next i
    Me.Caption = "DONE!"
End Sub

Private Sub Form_Load()
On Error GoTo Hell
    If Len(Command$) > 0 Then
        sFilePath = Command
        If FileLen(sFilePath) > 0 Then
            FileAndPath sFilePath, sFileName
            DoProcess
        End If
    Else
        ' set file association to .db
    End If
Exit Sub
Hell:
    MsgBox "ERROR: Invalid argument"
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(15) Then ' text = 1, url = 13, file = 15
        sFilePath = Data.Files(1)
        FileAndPath sFilePath, sFileName
        If sFileName = "thumbs.db" Then
            DoProcess
        Else
            Me.Caption = "Must be a thumbs.db file"
        End If
    End If
End Sub

Private Sub FileAndPath(ByRef sPath As String, Optional ByRef sFile As String)
    Dim i As Integer
    i = InStrRev(sPath, "\")
    sFile = Mid(sPath, i + 1)
    sPath = Mid(sPath, 1, i - 1)
End Sub

Private Function GetFileContents(ByVal sFile As String) As String
On Error Resume Next
    Dim iFile As Integer, i As Long
    iFile = FreeFile
    Open sFile For Binary As iFile
        GetFileContents = Space(LOF(1))
        Get #1, , GetFileContents
    Close iFile
End Function

Private Sub SetFileContents(ByVal sPath As String, ByVal sFile As String, ByVal contents As String)
    Dim iFile As Integer
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    iFile = FreeFile
    Open sPath & sFile For Output As iFile
        Print #iFile, contents;
    Close iFile
End Sub

Private Sub mnuAbout_Click()
    MsgBox "This program detects the beginning of a JPEG file inside the dragged file ends the JPEG at the point another one starts."
End Sub
