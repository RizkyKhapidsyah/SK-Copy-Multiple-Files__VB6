VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      Caption         =   "Copy Multiple files by Ali Ezzahir && Rizky Khapidsyah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4485
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

