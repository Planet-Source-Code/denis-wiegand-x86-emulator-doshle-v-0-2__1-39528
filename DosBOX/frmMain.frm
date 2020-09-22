VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DosHLE (DHLE) beta 0.01"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4245
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1920
      Left            =   6720
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   8
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   0
      Picture         =   "frmMain.frx":4954
      ScaleHeight     =   25
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   40.849
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin MSComDlg.CommonDialog CM 
         Left            =   3120
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MouseStickII"
            ImageKey        =   "MouseStickII"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2220
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":703F
            Key             =   "MouseStickII"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu tmp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEmu 
      Caption         =   "Emulator"
      Begin VB.Menu mnuSE 
         Caption         =   "Start Emulation"
      End
      Begin VB.Menu mnuSTOPEMU 
         Caption         =   "Stop Emualtion"
      End
      Begin VB.Menu mnuRC 
         Caption         =   "Reset CPU"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "?"
      Begin VB.Menu mnuA 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'8086 Emulator for old games
'Dont use new games like "DOOM" or "DarkForces" they wont work,
'because there`s no EMS emulation

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim I As Long
Dim ProgrammName As String

Private Sub mnuA_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuRC_Click()
brun = False
Emulate
End Sub

Private Sub mnuSTOPEMU_Click()
brun = False
Picture1.Cls
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "MouseStickII"
            'Zu erledigen: Schaltflächencode für 'MouseStickII' hinzufügen.
            mnuSE_Click
    End Select
End Sub


Private Sub Form_Load()
Me.Show
End Sub



Sub PrintDOS(ByVal code As Byte)
Dim wi As Single
Dim hi As Single
    x = (8 * (code Mod 16))
    y = (8 * Int(code \ 16))
    
    wi = 8
    hi = 8
    If code = 10 Then
        screenX = 0
        screenY = screenY + hi
        Exit Sub
    End If
    If code = 13 Then
        Exit Sub
    End If
    'Picture1.PaintPicture Picture2.Picture, 0, 0, wi, hi, x, y, 0.5, 0.5, vbSrcCopy
    'Picture1.Refresh
    BitBlt Picture1.hDC, screenX, screenY, wi, hi, Picture2.hDC, x, y, vbSrcCopy
    screenX = screenX + wi

    If screenX >= Picture1.Width - 10 Then
        screenX = 0
        screenY = screenY + hi
    End If
    Picture1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuOpen_Click()
With CM
    .FileName = ""
    .Filter = "COM files *.com|*.com"
    .ShowOpen
End With
If CM.FileName = "" Then Exit Sub
Open CM.FileName For Binary As #1
ProgrammName = CM.FileName
ReDim AsmText(LOF(1))
    Get #1, , AsmText
Close #1
For I = 0 To UBound(AsmText)
    Mem(CS + I) = AsmText(I)
Next I
End Sub

Private Sub mnuSE_Click()
If ProgrammName = "" Then
    MsgBox "Please load a COM File first", vbInformation, "Emualtor"
Else
    Picture1.Picture = LoadPicture("")
    Emulate
End If
End Sub



Private Sub Timer1_Timer()
StatusBar1.Panels(2).Text = ips
ips = 0
End Sub
