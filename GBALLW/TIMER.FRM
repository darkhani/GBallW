VERSION 4.00
Begin VB.Form MainFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "ÄÚ¹Í½Ã°è"
   ClientHeight    =   2280
   ClientLeft      =   2475
   ClientTop       =   1410
   ClientWidth     =   2220
   DrawStyle       =   2  'Dot
   Height          =   2685
   Left            =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   2220
   Top             =   1065
   Width           =   2340
   Begin VB.CommandButton Command2 
      Caption         =   "Ex"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2520
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   960
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Boon 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape SI 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   735
      Left            =   360
      Shape           =   2  'Oval
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "±¼¸²Ã¼"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   1080
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim C
Dim D
Dim E
Dim F
Private Sub Command1_Click()
AboutFrm.Show
End Sub

Private Sub Command2_Click()
End
End Sub

' Timer1 ÃÊ
' Timer2 ºÐ
' Timer3 ½Ã
Private Sub Form_Load()
Line1.BorderStyle = 0
MainFrm.Left = 0
MainFrm.Top = 0
End Sub
Private Sub Timer1_Timer()
a = Right(Time$, 2) 'ÃÊ
b = Val(a)
C = Mid(Time$, 4, 2) 'ºÐ
D = Val(C)
E = Left(Time$, 2) '½Ã
F = Val(E)
If F < 10 Then
  If D < 10 Then
    If b < 10 Then
      label1.Caption = 0 & F & " " & 0 & D & " " & 0 & b
      GoTo a:
     End If
    label1.Caption = 0 & F & " " & 0 & D & " " & b
    GoTo a
  End If
  If D > 10 And b < 10 Then
      label1.Caption = 0 & F & " " & D & " " & 0 & b
      GoTo a:
     End If

  label1.Caption = 0 & F & " " & D & " " & b
Else

If D < 10 Then
    If b < 10 Then
      label1.Caption = F & " " & 0 & D & " " & 0 & b
      GoTo a:
     End If
    label1.Caption = F & " " & 0 & D & " " & b
    GoTo a
  End If
  If D > 10 And b < 10 Then
      label1.Caption = F & " " & D & " " & 0 & b
      GoTo a:
     End If

End If
a:
Line1.BorderStyle = 1
If F > 12 Then
 F = F - 12
End If
If F >= 0 And F < 3 Then ' ¿ÞÂÊ ´«¾Ë ½Ã°£À» ³ªÅ¸³¿
  si.Left = 600: si.Top = 240
ElseIf F >= 3 And F < 6 Then
  si.Left = 600: si.Top = 480
ElseIf F >= 6 And F < 7 Then
  si.Left = 480: si.Top = 720
ElseIf F >= 7 And F < 9 Then
  si.Left = 360: si.Top = 600
ElseIf F >= 9 And F < 10 Then
  si.Left = 360: si.Top = 480
ElseIf F >= 10 And F < 12 Then
  si.Left = 360: si.Top = 360
End If

If D >= 0 And D < 15 Then '¿À¸¥ÂÊ ´«¾Ë ºÐÀ» ³ªÅ¸³¿
 boon.Left = 1560: boon.Top = 240

ElseIf D >= 15 And D < 20 Then
 boon.Left = 1560: boon.Top = 480

ElseIf D >= 20 And D < 30 Then
  boon.Left = 1560: boon.Top = 600

ElseIf D >= 30 And D < 35 Then
  boon.Left = 1440: boon.Top = 720

ElseIf D >= 35 And D < 40 Then
  boon.Left = 1320: boon.Top = 720

ElseIf D >= 40 And D < 45 Then
  boon.Left = 1320: boon.Top = 600
  
ElseIf D >= 45 And D < 55 Then
  boon.Left = 1320: boon.Top = 480

ElseIf D >= 55 And D < 59 Then
  boon.Left = 1320: boon.Top = 240
  
End If

If b = 0 Then
 Line1.X2 = 1080: Line1.Y2 = 240

ElseIf b >= 5 And b < 10 Then
 Line1.X2 = 1440: Line1.Y2 = 360

ElseIf b >= 10 And b < 15 Then
 Line1.X2 = 1800: Line1.Y2 = 720

ElseIf b >= 15 And b < 20 Then
 Line1.X2 = 1920: Line1.Y2 = 1080

ElseIf b >= 20 And b < 25 Then
 Line1.X2 = 1920: Line1.Y2 = 1560
 
ElseIf b >= 25 And b < 30 Then
 Line1.X2 = 1680: Line1.Y2 = 1800

ElseIf b >= 30 And b < 35 Then
 Line1.X2 = 1080: Line1.Y2 = 2040

ElseIf b >= 35 And b < 40 Then
 Line1.X2 = 720: Line1.Y2 = 1920

ElseIf b >= 40 And b < 45 Then
 Line1.X2 = 360: Line1.Y2 = 1560

ElseIf b >= 45 And b < 50 Then
 Line1.X2 = 240: Line1.Y2 = 1080

ElseIf b >= 50 And b < 55 Then
 Line1.X2 = 360: Line1.Y2 = 600

ElseIf b >= 55 And b < 59 Then
 Line1.X2 = 600: Line1.Y2 = 360

End If

End Sub



