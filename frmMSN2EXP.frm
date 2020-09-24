VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   960
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   1560
   Icon            =   "frmMSN2EXP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   960
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   120
   End
   Begin VB.PictureBox Pichook 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Menu mnuBar 
      Caption         =   "PopupMenu"
      Begin VB.Menu mnuMain1 
         Caption         =   "&Configure"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuMain2 
         Caption         =   "&Shut down"
         Index           =   0
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Private Sub Form_Load()
'I got this code from somewhere, I can't remember from where, but I did'nt write it, I just changed it
    t.cbSize = Len(t)
    t.hwnd = Pichook.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = "MSN Auto Nick" & Chr$(0) 'set the tooltiptext of the icon
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
    Form2.Visible = False
    
    Load FrmMain 'load the mainform
    FrmMain.Hide 'hide the mainform
    Call Timer1_Timer 'enable the "check for MSN" timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = Pichook.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Private Sub mnuMain1_Click(Index As Integer)
' the user clicks on "configure"
If FindMSN = "ok" Then 'if MSN is actif
    Load FrmMain 'load the mainform
    FrmMain.Show 'show the mainform
Else
    Call MsgBox("MSN is not active") 'inform the user
End If

End Sub

Private Sub mnuMain2_Click(Index As Integer)
'if the user clicked "exit"
FrmMain.cmdExit_Click 'click on the "exit" button on the mainform

End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Variant
    Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK: 'if the user doubleclicks with the left moeusebutton
                Call mnuMain1_Click(0)
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                Me.PopupMenu mnuBar, , , , mnuMain1(0) 'show the popupmenu with the first item in Bold
        End Select
        rec = False
    End If
End Sub

Private Function FindMSN()

intHwnd = FindWindow(vbNullString, "MSN Messenger Service") 'search the MSN window

If intHwnd = 0 Then 'if not found
    intHwnd = FindWindow(vbNullString, "MSN Messenger Service - AutoNick Enabled") 're-search it
    If intHwnd = 0 Then 'if not found
        FindMSN = "not" 'return "not found"
        Exit Function
    End If
End If
'if he found the window
FindMSN = "ok" 'return ok

End Function

Private Sub Timer1_Timer() 'the "check for MSN" timer

If FindMSN = "not" Then Exit Sub 'if MSN was not found, exit sub
'else
If FrmMain.lstNick.ListCount > 0 Then 'if there are any nicks in the list
    Call SetWindowText(intHwnd, "MSN Messenger Service - AutoNick Enabled") 'set the caption of the MSN window
    FrmMain.cmdApply_Click 'run the timer to set the nicks
End If
Timer1.Enabled = False 'disable this timer

End Sub
