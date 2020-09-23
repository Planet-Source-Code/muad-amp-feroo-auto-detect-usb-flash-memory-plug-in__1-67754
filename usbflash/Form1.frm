VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "usb flash  memory"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox plugin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1425
      Left            =   2880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1365
      ScaleWidth      =   1590
      TabIndex        =   1
      Top             =   2880
      Width           =   1650
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1365
      Left            =   2520
      Picture         =   "Form1.frx":7202
      ScaleHeight     =   1365
      ScaleWidth      =   1590
      TabIndex        =   0
      Top             =   480
      Width           =   1590
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim obj As New StdPicture
Private Declare Function GetDriveType Lib "kernel32" _
     Alias "GetDriveTypeA" _
    (ByVal nDrive As String) As Long

Dim fox(13)
Dim usb_drv As String

Private Sub Form_Load()
Dim c, i As Byte

For i = 68 To 80
c = c + 1
fox(c) = Chr(i) & ":"
Next
Set obj = pic

End Sub

Private Sub Timer1_Timer()

For i = 1 To 13
If GetDriveType(fox(i)) = 2 Then
pic.Picture = plugin.Picture
usb_drv = fox(i)
Exit Sub
Else
Set pic.Picture = obj
usb_drv = ""
End If
Next
End Sub

