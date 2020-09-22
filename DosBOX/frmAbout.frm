VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About DosHLE"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Command1 
      Caption         =   "Okey"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1275
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1215
      ScaleWidth      =   4140
      TabIndex        =   0
      Top             =   120
      Width           =   4200
   End
   Begin VB.Label Label1 
      Caption         =   "DosHLE by Denis Wiegand 2002 (c)  http://www.geocities.com/deniswi2002"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
