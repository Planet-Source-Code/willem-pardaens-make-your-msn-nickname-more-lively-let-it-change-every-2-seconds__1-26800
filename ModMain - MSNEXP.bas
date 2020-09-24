Attribute VB_Name = "ModMain"
Option Explicit

'First Make a Reference to the MSN application (.exe)
Public MSN As Messenger.MsgrObject 'the MSN object

Public Dirslash As String
Public strTemp As String
Public ISNickNumber As Integer
Public intTemp As Integer
Public intHwnd As Integer
Public lngTemp As Long
Public arrNick() As String

'the apis
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Sub Main()

Select Case Right(App.Path, 1) 'If the path ends with a backslash
Case "\"
    Dirslash = App.Path 'add nothing
Case Else
    Dirslash = App.Path & "\" ' add a backslash
End Select

Set MSN = New Messenger.MsgrObject 'create new instance of the MSN object

End Sub
