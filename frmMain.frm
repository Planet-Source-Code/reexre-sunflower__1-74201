VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PHI"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chJpg 
      Caption         =   "Save Frames"
      Height          =   255
      Left            =   13200
      TabIndex        =   9
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   " phi"
      Height          =   615
      Left            =   13200
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox chRad 
      Caption         =   "Radius"
      Height          =   255
      Left            =   13200
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   12960
      ScaleHeight     =   1215
      ScaleWidth      =   1935
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.HScrollBar sSpeed 
      Height          =   375
      Left            =   13200
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   3480
      Value           =   100
      Width           =   975
   End
   Begin VB.TextBox txtN 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   13200
      TabIndex        =   3
      Text            =   "500"
      ToolTipText     =   "N of Points"
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   " stop"
      Height          =   615
      Left            =   13200
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7095
      ScaleWidth      =   9015
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "speed"
      Height          =   255
      Left            =   13200
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   13200
      TabIndex        =   5
      Top             =   7080
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command1_Click()
    Command1.Enabled = False
    Animation PIC, Val(txtN), sSpeed.Value

End Sub

Private Sub Command2_Click()
    ExitLoop = True

End Sub

Private Sub Form_Load()


    PIC.AutoSize = True
    PIC = LoadPicture(App.Path & "\sunflower.jpg")

    Picture1.Width = PIC.Width * 0.33
    Picture1.Height = PIC.Height * 0.33
    SetStretchBltMode Picture1.hDC, STRETCHMODE
    StretchBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, PIC.hDC, _
               0, 0, PIC.Width, PIC.Height, vbSrcCopy



    PIC.Cls
    PIC.AutoSize = False
    PIC.Refresh



    PIC.Height = 640
    PIC.Width = Int(4 / 3 * PIC.Height)

    cX = PIC.Width \ 2
    cY = PIC.Height \ 2

    Me.Show
    Me.Refresh
    DoEvents

    If Dir(App.Path & "\Frames", vbDirectory) = "" Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.jpg") <> "" Then Kill App.Path & "\Frames\*.jpg"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase P
    End

End Sub

