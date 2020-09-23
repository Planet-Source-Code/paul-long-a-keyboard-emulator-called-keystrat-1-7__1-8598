VERSION 5.00
Begin VB.Form Info 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton About_Ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "Info.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "K E Y S T R A T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "pmdlong@ntlworld.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Or send an e-mail to"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "http://homepage.ntlworld.com/pmdlong/keystrat.htm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "If you require any further information about the program or the aims of the project please visit the project website website."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "This program was developed to help evaluate possible new choice strategies for use in keyboard emulators."
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "This program was developed as part of a final year Computer Systems project at Coventry University."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "By  Paul Long"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Keyboard Emulator for Choice Strategy Evaluation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem: KEYSTRAT 1.7 - Info.frm
Rem: ***********************
Option Explicit
Rem: Unload Form On OK Click
Private Sub About_Ok_Click()
    Unload Info
End Sub

