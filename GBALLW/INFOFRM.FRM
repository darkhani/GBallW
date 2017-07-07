VERSION 4.00
Begin VB.Form InfoFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4365
   ClientLeft      =   2205
   ClientTop       =   1530
   ClientWidth     =   4470
   Height          =   4770
   Left            =   2145
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4470
   Top             =   1185
   Width           =   4590
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4560
      Top             =   1560
   End
   Begin VB.PictureBox Picture1 
      Height          =   4500
      Left            =   0
      Picture         =   "INFOFRM.frx":0000
      ScaleHeight     =   4470
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   -120
      Width           =   4455
   End
End
Attribute VB_Name = "InfoFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
InfoFrm.Hide
Timer1.Interval = 0
MainFrm.Show
End Sub


