VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   600
      ScaleHeight     =   1275
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   600
      Picture         =   "SmartStretch.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set Picture2.Picture = SmartStretch(Picture1.Picture, ScaleY(Picture1.Picture.Height), ScaleX(Picture1.Picture.Width), Picture2.ScaleHeight, Picture2.ScaleWidth)
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture("c:\windows\clouds.bmp")
End Sub

Public Function SmartStretch(ByVal Picture As IPictureDisp, ByVal Height As Long, ByVal Width As Long, ByVal NHeight As Long, ByVal NWidth As Long) As IPictureDisp
Me.Cls
If NWidth < (Width * (NHeight / Height)) Then
PaintPicture Picture, 1, 1, NWidth, (Height * (NWidth / Width))
Else
PaintPicture Picture, 1, 1, (Width * (NHeight / Height)), NHeight
End If

Set SmartStretch = Me.Image
Me.Cls
End Function

