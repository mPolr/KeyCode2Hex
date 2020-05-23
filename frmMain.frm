VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KeyCode -> HEX"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKeyCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtHEX 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtKey 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblAbout 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "О программе"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   615
      MouseIcon       =   "frmMain.frx":42B2
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      MouseIcon       =   "frmMain.frx":4B7C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Код:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hex-значение:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Клавиша:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Sub Form_Initialize()
 InitCommonControls
End Sub
Private Sub lblAbout_Click()
 MsgBox "KeyCode2Hex v" & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & "(c) 2005 by mPolr" & vbNewLine & " " & vbNewLine & "http://mpolr.narod.ru/", vbInformation, "О программе..."
End Sub
Private Sub lblExit_Click()
 End
End Sub
Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
txtKey.Text = ""
Select Case KeyCode
 Case 8
  txtKey.Text = "[BACKSPACE]"
 Case 9
  txtKey.Text = "[TAB]"
 Case 13
  txtKey.Text = "[ENTER]"
 Case 16
  txtKey.Text = "[SHIFT]"
 Case 17
  txtKey.Text = "[CONTROL]"
 Case 18
  txtKey.Text = "[ALT]"
 Case 20
  txtKey.Text = "[CAPS LOCK]"
 Case 27
  txtKey.Text = "[ESC]"
 Case 32
  txtKey.Text = "[SPACE]"
 Case 33
  txtKey.Text = "[PAGE UP]"
 Case 34
  txtKey.Text = "[PAGE DOWN]"
 Case 35
  txtKey.Text = "[END]"
 Case 36
  txtKey.Text = "[HOME]"
 Case 37
  txtKey.Text = "[LEFT ARROW]"
 Case 38
  txtKey.Text = "[UP ARROW]"
 Case 39
  txtKey.Text = "[RIGHT ARROW]"
 Case 40
  txtKey.Text = "[DOWN ARROW]"
 Case 45
  txtKey.Text = "[INSERT]"
 Case 91
  txtKey.Text = "[WINDOWS]"
 Case 112
  txtKey.Text = "[F1]"
 Case 113
  txtKey.Text = "[F2]"
 Case 114
  txtKey.Text = "[F3]"
 Case 115
  txtKey.Text = "[F4]"
 Case 116
  txtKey.Text = "[F5]"
 Case 117
  txtKey.Text = "[F6]"
 Case 118
  txtKey.Text = "[F7]"
 Case 119
  txtKey.Text = "[F8]"
 Case 120
  txtKey.Text = "[F9]"
 Case 121
  txtKey.Text = "[F10]"
 Case 122
  txtKey.Text = "[F11]"
 Case 123
  txtKey.Text = "[F12]"
 Case 144
  txtKey.Text = "[NUM LOCK]"
 Case 145
  txtKey.Text = "[SCROLL LOCK]"
 Case 166
  txtKey.Text = "[PREVIOUS PAGE]"
 Case 167
  txtKey.Text = "[NEXT PAGE]"
 Case 168
  txtKey.Text = "[REFRESH]"
 Case 169
  txtKey.Text = "[STOP]"
 Case 170
  txtKey.Text = "[SEARCH]"
 Case 171
  txtKey.Text = "[FAVORITES]"
 Case 172
  txtKey.Text = "[BROWSER]"
 Case 180
  txtKey.Text = "[MAIL]"
End Select
txtKeyCode.Text = KeyCode
txtHEX.Text = "&H" & Hex(KeyCode)
End Sub
