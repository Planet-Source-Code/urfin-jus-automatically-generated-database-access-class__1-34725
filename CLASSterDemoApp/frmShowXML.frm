VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmShowXML 
   Caption         =   "XML View"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser xmlViewer 
      Height          =   6555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10710
      ExtentX         =   18891
      ExtentY         =   11562
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmShowXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Let xml(ByVal Msg As String)
    Dim FileName As String
    On Error GoTo errHandler
    If Not Me.Visible Then Me.Show
    While xmlViewer.Busy
        DoEvents
    Wend
    FileName = App.Path & "\" & "Temp.xml"
    SaveTempFile FileName, Msg
    xmlViewer.Navigate FileName
    Exit Property
errHandler:
    MsgBox "Fatal error: " & Err.Description
End Property

Private Sub SaveTempFile(ByVal FileName As String, ByVal Msg As String)
    Dim S As ADODB.Stream
    On Error GoTo errHandler
    Set S = New ADODB.Stream
    S.Open
    S.WriteText Msg
    S.SaveToFile FileName, adSaveCreateOverWrite
    S.Close
    Exit Sub
errHandler:
    MsgBox "Fatal error when saving temp file in frm0XML: " & Err.Description
End Sub

Private Sub Form_Resize()
  xmlViewer.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


