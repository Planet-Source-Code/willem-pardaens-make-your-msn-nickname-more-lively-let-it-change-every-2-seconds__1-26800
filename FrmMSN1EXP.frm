VERSION 5.00
Object = "{3489755E-DC13-11D4-9242-000102711081}#8.0#0"; "METALCBPROJ.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WPsoftware - MSNAutoNick"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "FrmMSN1EXP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrShowNick 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   720
   End
   Begin MetalCBProj.MetalCB cmdReset 
      Height          =   300
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      FontSize        =   8.25
      FontCharset     =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontStrike      =   0   'False
      FontUnder       =   0   'False
      FontWeight      =   400
      Caption         =   "&Reset"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Red             =   200
      Green           =   0
      Blue            =   0
      Multiplier      =   -2
   End
   Begin VB.Frame frameSpeed 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   2450
      Begin MSComCtl2.UpDown UDSpeed 
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "lblSpeed2"
         BuddyDispid     =   196630
         OrigLeft        =   1920
         OrigTop         =   480
         OrigRight       =   2160
         OrigBottom      =   855
         Max             =   120
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MetalCBProj.MetalCB cmdNIU2 
         Height          =   300
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         FontSize        =   8.25
         FontCharset     =   0
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontStrike      =   0   'False
         FontUnder       =   0   'False
         FontWeight      =   400
         Caption         =   "&Change speed"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Red             =   255
         Green           =   255
         Multiplier      =   -2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Speed in seconds:"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblSpeed2 
         Caption         =   "2"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3.5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   680
         Width           =   1720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   1080
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2400
         Y1              =   1080
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   2400
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin MetalCBProj.MetalCB cmdChange 
      Height          =   300
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      FontSize        =   8.25
      FontCharset     =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontStrike      =   0   'False
      FontUnder       =   0   'False
      FontWeight      =   400
      Caption         =   "Change your MSN Nickname"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Blue            =   100
      Multiplier      =   -2
   End
   Begin VB.Frame frameChange 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2450
      Begin VB.TextBox txtNick 
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin MetalCBProj.MetalCB cmdNIU 
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         FontSize        =   8.25
         FontCharset     =   0
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontStrike      =   0   'False
         FontUnder       =   0   'False
         FontWeight      =   400
         Caption         =   "&Change your MSN nickname"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Green           =   255
         Blue            =   255
         Multiplier      =   -2
         Enabled         =   0
      End
      Begin MetalCBProj.MetalCB cmdAdd 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         FontSize        =   8.25
         FontCharset     =   0
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontStrike      =   0   'False
         FontUnder       =   0   'False
         FontWeight      =   400
         Caption         =   "&Add"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Green           =   100
         Multiplier      =   -2
      End
      Begin MetalCBProj.MetalCB cmdRemove 
         Height          =   300
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         FontSize        =   8.25
         FontCharset     =   0
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontStrike      =   0   'False
         FontUnder       =   0   'False
         FontWeight      =   400
         Caption         =   "&Remove"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Red             =   100
         Multiplier      =   -2
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   0
         X2              =   2400
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   10
         X2              =   10
         Y1              =   120
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2400
         Y1              =   240
         Y2              =   1560
      End
   End
   Begin MetalCBProj.MetalCB cmdApply 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      FontSize        =   8.25
      FontCharset     =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontStrike      =   0   'False
      FontUnder       =   0   'False
      FontWeight      =   400
      Caption         =   "&Apply"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Red             =   100
      Blue            =   100
      Multiplier      =   -2
   End
   Begin MetalCBProj.MetalCB cmdHide 
      Height          =   300
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      FontSize        =   8.25
      FontCharset     =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontStrike      =   0   'False
      FontUnder       =   0   'False
      FontWeight      =   400
      Caption         =   "&Hide"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Red             =   255
      Green           =   100
      Blue            =   0
      Multiplier      =   -2
   End
   Begin VB.ListBox lstNick 
      Height          =   3570
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin MetalCBProj.MetalCB cmdExit 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      FontSize        =   8.25
      FontCharset     =   0
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontStrike      =   0   'False
      FontUnder       =   0   'False
      FontWeight      =   400
      Caption         =   "&Exit"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Red             =   255
      Multiplier      =   -2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright(c) 2001 by WPsoftware"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   17
      Top             =   4560
      Width           =   2370
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   4920
      Picture         =   "FrmMSN1EXP.frx":0442
      Stretch         =   -1  'True
      Top             =   80
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   1560
      Picture         =   "FrmMSN1EXP.frx":0D4D
      Stretch         =   -1  'True
      Top             =   80
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      X1              =   240
      X2              =   6720
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################'
'##                                          ##'
'##     WPsoftware - MSN Auto Nick v.1       ##'
'##     By Willem Pardaens                   ##'
'##     Copyright(c) 2001 by WPsoftware      ##'
'##     This code is freeware, but please    ##'
'##        vote for me at PSC and add me     ##'
'##        to your MSN contact list.         ##'
'##        hotmail adress:                   ##'
'##        willem_pardaens@hotmail.com       ##'
'##     When you have questions or suggest   ##'
'##        ions, mail me:                    ##'
'##        helpdesk.wpsoftware@tijd.com      ##'
'##                                          ##'
'##     HAVE FUN!!, WP                       ##'
'##                                          ##'
'##############################################'

Private Sub cmdHide_Click()

Me.Hide 'Hide the main form

End Sub

Private Sub cmdReset_Click()

If MsgBox("Are you sure you want to delete the current nickname settings?", vbInformation + vbYesNo) = vbNo Then Exit Sub 'ask user if he's sure about it
txtNick.Text = "" 'clear textbox
lstNick.Clear 'clear the list

End Sub

Private Sub Form_Load()

Call ModMain.Main 'Do main stuff (initializing)

Me.Height = 1000 'set the height of the form

Line1.X1 = -100 'set the line's left
Line1.X2 = Me.Width + 100 'set the line's width

intTemp = RegGet("Software\WPsoftware\MSNAutoNick", "Speed") 'Get the speed value from the reg
If intTemp <> 0 Then 'if it is not empty
    UDSpeed.Value = intTemp 'fill the UpDowncontrol with the value
Else
    UDSpeed.Value = 10 'fill the updown with the default value
End If

lstNick.Clear 'clear the lsit

Call LoadFromFile 'load the list of nicks from file

End Sub

Private Sub cmdAdd_Click()

If txtNick.Text = "" Then MsgBox ("You have to fill in a nick"): Exit Sub 'check if the user has entered a value

lstNick.AddItem txtNick.Text, lstNick.ListIndex + 1 'add nick to list
lstNick.ListIndex = lstNick.ListCount - 1 'set listindex at the last item
If lstNick.ListCount = 30 Then MsgBox ("You reached the maximum of 30 nicks") 'set the maximum at 30
txtNick.SelStart = 0 'select the nick in the textbox
txtNick.SelLength = Len(txtNick) '=
txtNick.SetFocus 'set the focus on the textbox

End Sub

Public Sub cmdApply_Click()

If lstNick.ListCount = 0 Then MsgBox ("There are no nicks"): Exit Sub 'check if there are any nicks
If lstNick.ListCount > 30 Then MsgBox ("Too many nicks"): Exit Sub 'check if there arent too many nicks

txtNick.BackColor = &HC0C0C0 'make it gray
lstNick.BackColor = &HC0C0C0 '=
txtNick.Locked = True 'lock it
cmdAdd.Enabled = 0 'disable
cmdRemove.Enabled = 0 'disable
frameSpeed.Enabled = False 'disable

ReDim arrNick(lstNick.ListCount - 1) As String 'redim the array
For intTemp = 1 To lstNick.ListCount 'loop through the listbox
    lstNick.ListIndex = intTemp - 1
    arrNick(intTemp - 1) = lstNick.Text 'add the value to the array
    DoEvents
Next intTemp 'reloop
ISNickNumber = 0 'set the value to the beginning of the array
tmrShowNick.Interval = lblSpeed2.Caption * 500 'set the timer interval
tmrShowNick.Enabled = True 'enable the timer
Me.Height = 1000 'set the height
Me.Hide 'hide the mainform

Call SaveToFile 'save the settings

End Sub

Public Sub cmdExit_Click()

Call SaveToFile 'save the settings again :) (te be sure)
If Not intHwnd = 0 Then 'if MSN is actif
    Call SetWindowText(intHwnd, "MSN Messenger Service") 'set the caption to normal
End If

Call RegPut("Software\WPsoftware\MSNAutoNick", "Speed", lblSpeed2.Caption) 'save the speed in mem
Set MSN = Nothing 'set MSN to nothing

For Each Form In Forms 'unload forms
    Unload Form
Next Form
End 'end the program

End Sub

Private Sub cmdChange_Click()

txtNick.BackColor = vbWhite 'make it white
lstNick.BackColor = vbWhite 'make it white
txtNick.Locked = False 'delock it
cmdAdd.Enabled = 1 'enable
cmdRemove.Enabled = 1 'enable
frameSpeed.Enabled = True 'enable
Me.Height = 5200 'set the height
txtNick.SetFocus 'set focus on the textbox
tmrShowNick.Enabled = False 'disable the timer 'cause you're in "design" mode

End Sub

Private Sub cmdRemove_Click()

lstNick.RemoveItem lstNick.ListIndex 'remove the item from the list
cmdRemove.Enabled = 0 'disable the remove commandbutton
lstNick.ListIndex = lstNick.ListCount - 1 'set the listindex to the last item

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then Cancel = 1 'cancel the unload if the user clicked on the X on the controlbox

End Sub

Private Sub Form_Resize()

Me.Left = Screen.Width - Me.Width - 50 'resize a bit
Me.Top = Screen.Height - Me.Height - 1000 'resize a bit

End Sub

Private Sub Image1_Click()

Call MsgBox("Copyright(c) 2001 by WPsoftware." & vbCrLf & vbCrLf & "For help mailto: helpdesk.wpsoftware@tijd.com", vbInformation + vbOKOnly, "about")

End Sub

Private Sub Image2_Click()

Call MsgBox("Copyright(c) 2001 by WPsoftware." & vbCrLf & vbCrLf & "For help mailto: helpdesk.wpsoftware@tijd.com", vbInformation + vbOKOnly, "about")

End Sub

Private Sub lblSpeed2_Change()

lblSpeed.Caption = (lblSpeed2.Caption / 2) 'Because the updowncontrol can't work with a half, i make it a half

End Sub

Private Sub lstNick_Click()

On Error Resume Next
If lstNick.ListIndex = -1 Then cmdRemove.Enabled = 0 Else cmdRemove.Enabled = 1 'check if the user selected a nick
txtNick.Text = lstNick.Text 'set the listitem in the textbox
txtNick.SelStart = 0 ' select the nick
txtNick.SelLength = Len(txtNick) ' select the nick
txtNick.SetFocus 'set focus on the textbox

End Sub

Private Sub lstNick_KeyPress(KeyAscii As Integer)

If KeyAscii = 100 Then
    KeyAscii = 0
    cmdRemove_Click 'if the use pressed "d" delete the listitem
End If

End Sub

Private Sub tmrShowNick_Timer()

Call SetNick(arrNick(ISNickNumber)) 'call the SetNick procedure with the item in the array as a parameter

If ISNickNumber < UBound(arrNick) Then 'if the nicknumber is not the last one
    ISNickNumber = ISNickNumber + 1 'increase by 1
Else
    ISNickNumber = 0 'set to 0 again
End If

End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    cmdAdd_Click 'if the user pressed enter, add the item to the list
End If

End Sub

Private Sub SetNick(strTemp)

On Error GoTo ErrH 'there would be an error if MSN is not active

'this is the actual line:
MSN.Services(0).FriendlyName = strTemp
'in the mainmodule we set MSN as an messenger object
'here we use the object to change the "friendlyname" or nickname to the parameter value

DoEvents

Exit Sub
ErrH:   tmrShowNick.Enabled = False 'disable the timer
        Call SetWindowText(intHwnd, "MSN Messenger Service") 'reset the caption of the MSN window
        Form2.Timer1.Enabled = True 'enable the "check for MSN" timer
End Sub

Private Sub LoadFromFile()

strTemp = RegGet("Software\WPsoftware\MSNAutoNick", "File") 'get the filename from the reg
If strTemp = "" Then 'if the value is "", fill it with the default
    strTemp = Dirslash & "Nicks.dat"
    Open strTemp For Output As #1 'create the file
    Close #1
    Call RegPut("Software\WPsoftware\MSNAutoNick", "File", Dirslash & "Nicks.dat") 'save the filename
End If

Open strTemp For Input As #1 'open the file to read from it
Do While Not EOF(1) 'if the file isnt empty
    Line Input #1, strTemp
    lstNick.AddItem strTemp 'add the nick to the list
    DoEvents
Loop
Close #1 'close the file

End Sub

Private Sub SaveToFile()

strTemp = RegGet("Software\WPsoftware\MSNAutoNick", "File") 'get the filename from the reg
If strTemp = "" Then 'if the value is "", fill it with the default
    strTemp = Dirslash & "Nicks.dat"
    Call RegPut("Software\WPsoftware\MSNAutoNick", "File", Dirslash & "Nicks.dat") 'save the filename
End If

Open strTemp For Output As #1 'open the file to write in it
For intTemp = 0 To lstNick.ListCount - 1 'loop through the list
    lstNick.ListIndex = intTemp
    Print #1, lstNick.Text 'write the nick in the file
    DoEvents
Next intTemp
Close #1 'close the file

End Sub
