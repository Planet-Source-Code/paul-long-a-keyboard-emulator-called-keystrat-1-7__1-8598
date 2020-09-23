VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KeyStrat 1.7"
   ClientHeight    =   6330
   ClientLeft      =   1770
   ClientTop       =   2325
   ClientWidth     =   7725
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7725
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   5640
      Top             =   360
   End
   Begin VB.PictureBox Caps_Light 
      BackColor       =   &H8000000B&
      Height          =   135
      Left            =   6000
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   59
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox Shift_Light 
      BackColor       =   &H8000000B&
      Height          =   135
      Left            =   6720
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   58
      Top             =   600
      Width           =   375
   End
   Begin VB.HScrollBar Speed_Bar 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   1
      Min             =   200
      TabIndex        =   57
      Top             =   480
      Value           =   100
      Width           =   5415
   End
   Begin VB.TextBox Keyboard_Text 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      Top             =   3600
      Width           =   7245
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "stop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   55
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "space"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   54
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2880
      Width           =   6255
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   52
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   51
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   49
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   48
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   46
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   45
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   44
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   43
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   42
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "shift"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "enter"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "lock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "["
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "tab"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "bksp"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Keyboard_Array 
      BackColor       =   &H80000016&
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Speed_Bar_Speed 
      BackColor       =   &H8000000B&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   63
      Top             =   240
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      DrawMode        =   6  'Mask Pen Not
      X1              =   -240
      X2              =   8280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Shift_Light_Label 
      BackColor       =   &H8000000B&
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   62
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Speed_Bar_Label 
      BackColor       =   &H8000000B&
      Caption         =   "Speed ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   61
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Caps_Light_Label 
      BackColor       =   &H8000000B&
      Caption         =   "Caps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   60
      Top             =   360
      Width           =   375
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Start 
         Caption         =   "Start"
      End
      Begin VB.Menu Stop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Input 
      Caption         =   "Input"
      Begin VB.Menu Switch 
         Caption         =   "Simulated Switch Press"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu Strategy 
      Caption         =   "Strategy"
      Begin VB.Menu Strategy1 
         Caption         =   "Strategy1"
         Checked         =   -1  'True
      End
      Begin VB.Menu Strategy2 
         Caption         =   "Strategy2"
      End
      Begin VB.Menu Strategy3 
         Caption         =   "Strategy3"
      End
      Begin VB.Menu Strategy4 
         Caption         =   "Strategy4"
      End
   End
   Begin VB.Menu Layout 
      Caption         =   "Layout"
      Begin VB.Menu Layout1 
         Caption         =   "Qwerty"
         Checked         =   -1  'True
      End
      Begin VB.Menu Layout2 
         Caption         =   "Dvorak"
      End
      Begin VB.Menu Layout3 
         Caption         =   "Alphabetic"
      End
   End
   Begin VB.Menu Sound 
      Caption         =   "Sound"
      Begin VB.Menu Sound1 
         Caption         =   "Motion"
         Checked         =   -1  'True
      End
      Begin VB.Menu Sound2 
         Caption         =   "Keypress"
      End
      Begin VB.Menu Sound3 
         Caption         =   "No Sound"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Using 
         Caption         =   "Using KeyStrat"
      End
      Begin VB.Menu About 
         Caption         =   "About KeyStrat"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem: KEYSTRAT 1.7 - Main.frm
Rem: ***********************
Option Explicit
Rem: Constant Values
Const BKSP_KEY = 13
Const TAB_KEY = 14
Const LOCK_KEY = 28
Const ENTER_KEY = 40
Const SHIFT_KEY = 41
Const CLEAR_KEY = 53
Const SPACE_KEY = 54
Const STOP_KEY = 55
Rem: Declare Variables
Dim doFlag As Boolean
Dim i As Integer
Rem: Program Setup
Private Sub Form_Load()
    doFlag = False 'Strategy is Off
    Speed_Bar.Value = 101 'Sets Default Speed
    Layout1_Setup 'Loads Default Layout (Qwerty)
End Sub
Rem: Menu | File | Start
Private Sub Start_Click()
    doFlag = True
End Sub
Rem: Menu | File | Stop
Private Sub Stop_Click()
    doFlag = False
    Clear_Keyboard
End Sub
Rem: Menu | File | Exit
Private Sub Exit_Click()
    Unload Main
End Sub
Rem: Menu | Input | Simulated Switch Press (F1)
Private Sub Switch_Click()
    If Strategy1.Checked And doFlag Then
        Strategy1_Inc_Stage 'Each Keypress Moves Strategy Into Next Stage
    ElseIf Strategy2.Checked And doFlag Then
        Strategy2_Inc_Stage 'Each Keypress Moves Strategy Into Next Stage
    ElseIf Strategy3.Checked And doFlag Then
        Strategy3_Inc_Stage 'Each Keypress Moves Strategy Into Next Stage
    ElseIf Strategy4.Checked And doFlag Then
        Strategy4_Inc_Stage 'Each Keypress Moves Strategy Into Next Stage
    End If
End Sub
Rem: Menu | Strategy | Strategy1
Private Sub Strategy1_Click()
    Strategy_Switch 'Resets Current Strategy
    Strategy1.Checked = True
    Strategy2.Checked = False
    Strategy3.Checked = False
    Strategy4.Checked = False
End Sub
Rem: Menu | Strategy | Strategy2
Private Sub Strategy2_Click()
    Strategy_Switch 'Resets Current Strategy
    Strategy1.Checked = False
    Strategy2.Checked = True
    Strategy3.Checked = False
    Strategy4.Checked = False
End Sub
Rem: Menu | Strategy | Strategy3
Private Sub Strategy3_Click()
    Strategy_Switch 'Resets Current Strategy
    Strategy1.Checked = False
    Strategy2.Checked = False
    Strategy3.Checked = True
    Strategy4.Checked = False
End Sub
Rem: Menu | Strategy | Strategy4
Private Sub Strategy4_Click()
    Strategy_Switch 'Resets Current Strategy
    Strategy1.Checked = False
    Strategy2.Checked = False
    Strategy3.Checked = False
    Strategy4.Checked = True
End Sub
Rem: Menu | Layout | Layout1 (Qwerty - Default Setting)
Private Sub Layout1_Click()
    Layout1.Checked = True
    Layout2.Checked = False
    Layout3.Checked = False
    Layout1_Setup 'Relabel Keyboard with Qwerty Layout
End Sub
Rem: Menu | Layout | Layout2 (Dvorak)
Private Sub Layout2_Click()
    Layout1.Checked = False
    Layout2.Checked = True
    Layout3.Checked = False
    Layout2_Setup 'Relabel Keyboard with Dvorak Layout
End Sub
Rem: Menu | Layout | Layout3 (Alphabetic)
Private Sub Layout3_Click()
    Layout1.Checked = False
    Layout2.Checked = False
    Layout3.Checked = True
    Layout3_Setup 'Relabel Keyboard with Alphabetic Layout
End Sub
Rem: Menu | Sound | Motion (Default Setting)
Private Sub Sound1_Click()
    Sound1.Checked = True
    Sound2.Checked = False
    Sound3.Checked = False
End Sub
Rem: Menu | Sound | Keypress
Private Sub Sound2_Click()
    Sound1.Checked = False
    Sound2.Checked = True
    Sound3.Checked = False
End Sub
Rem: Menu | Sound | No Sound
Private Sub Sound3_Click()
    Sound1.Checked = False
    Sound2.Checked = False
    Sound3.Checked = True
End Sub
Rem: Menu | Help | Using KeyStrat
Private Sub Using_Click()
    Dim lRet As Long
    lRet = Shell("winhelp.exe helpfile.hlp", 1) 'Load Help File
End Sub
Rem: Menu | Help | About KeyStrat
Private Sub About_Click()
    Info.Show 'Show About Dialogue Box
End Sub
Rem: Feature | Timer
Private Sub Timer_Timer()
    If Strategy1.Checked And doFlag Then
        Strategy1_Move
    ElseIf Strategy2.Checked And doFlag Then
        Strategy2_Move
    ElseIf Strategy3.Checked And doFlag Then
        Strategy3_Move
    ElseIf Strategy4.Checked And doFlag Then
        Strategy4_Move
    End If
    If Sound1.Checked And doFlag Then
        Beep 'Beeps With Timer
    End If
End Sub
Rem: Handle Changes to Speed Bar
Private Sub Speed_Bar_Change()
    Timer.Interval = Speed_Bar.Value * 10
    Speed_Bar_Speed.Caption = 201 - Speed_Bar.Value
End Sub
Rem: Toggle Shift Indicator
Private Sub Shift_Light_Switch()
    If (Shift_Light.BackColor = &HFF00&) Then
        Shift_Light.BackColor = &H8000000B
    ElseIf (Shift_Light.BackColor = &H8000000B) Then
        Shift_Light.BackColor = &HFF00&
    End If
End Sub
Rem: Toggle Lock Indicator
Private Sub Caps_Light_Switch()
    If (Caps_Light.BackColor = &HFF00&) Then
        Caps_Light.BackColor = &H8000000B
    ElseIf (Caps_Light.BackColor = &H8000000B) Then
        Caps_Light.BackColor = &HFF00&
    End If
End Sub
Rem: Turn Keypress into Textbox Output
Public Sub Keyboard_Array_Click(Index As Integer)
    Dim text_length As Integer
    Dim num_chars As Integer
    Dim last_char As String
    If (Index = 7) And (Shift_Light.BackColor = &HFF00&) Then
            'Can not simply copy caption since this reads "&&" (see layout files)
            Keyboard_Text.Text = Keyboard_Text.Text + "&"
    ElseIf (Index = BKSP_KEY) Then 'bksp
            If (Len(Keyboard_Text.Text) > 0) Then
            text_length = Len(Keyboard_Text.Text)
            last_char = Right(Keyboard_Text.Text, (text_length - (text_length - 1)))
            If last_char = vbLf Then
                num_chars = 2
            Else
                num_chars = 1
            End If
            Keyboard_Text.Text = Left(Keyboard_Text.Text, (text_length - num_chars))
        End If
    ElseIf (Index = TAB_KEY) Then 'tab
        Keyboard_Text.Text = Keyboard_Text.Text + vbTab
    ElseIf (Index = LOCK_KEY) Then 'lock
        Caps_Light_Switch
        If Layout1.Checked Then
            Layout1_Toggle_Lock
        ElseIf Layout2.Checked Then
            Layout2_Toggle_Lock
        ElseIf Layout3.Checked Then
            Layout3_Toggle_Lock
        End If
    ElseIf (Index = ENTER_KEY) Then 'enter
        Keyboard_Text.Text = Keyboard_Text.Text + vbCrLf
    ElseIf (Index = SHIFT_KEY) Then 'shift
        Shift_Light_Switch
        If Layout1.Checked Then
            Layout1_Toggle_Shift
        ElseIf Layout2.Checked Then
            Layout2_Toggle_Shift
        ElseIf Layout3.Checked Then
            Layout3_Toggle_Shift
        End If
    ElseIf (Index = CLEAR_KEY) Then 'clear
        Keyboard_Text.Text = ""
    ElseIf (Index = SPACE_KEY) Then 'space
        Keyboard_Text.Text = Keyboard_Text.Text + " "
    ElseIf (Index = STOP_KEY) Then 'stop
        Strategy_Switch
        Stop_Click
    Else 'letter, number or symbol
        Keyboard_Text.Text = Keyboard_Text.Text + Keyboard_Array(Index).Caption
    End If
    Keyboard_Text.SelStart = Len(Keyboard_Text.Text)
    Keyboard_Text.SetFocus
    If Sound2.Checked And doFlag Then
        Beep 'Beeps when key is selected
    End If
End Sub
Rem: Clear All Highlights Off Keyboard
Public Sub Clear_Keyboard()
    For i = 0 To 55
        Keyboard_Array(i).BackColor = &H80000016
    Next i
End Sub
Rem: Reset Current Strategy Before Starting Another
Private Sub Strategy_Switch()
    Clear_Keyboard 'Clear All Highlights Off Keyboard
    Keyboard_Text.Text = "" 'Clear Text Box
    If Strategy1.Checked Then
        Strategy1_Reset 'Reset Strategy1 Variables
    ElseIf Strategy2.Checked Then
        Strategy2_Reset 'Reset Strategy2 Variables
    ElseIf Strategy3.Checked Then
        Strategy3_Reset 'Reset Strategy3 Variables
    ElseIf Strategy4.Checked Then
        Strategy4_Reset 'Reset Strategy4 Variables
    End If
End Sub
