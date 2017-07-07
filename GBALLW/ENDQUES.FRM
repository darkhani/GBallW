VERSION 4.00
Begin VB.Form EndQues 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "끝냅니까?"
   ClientHeight    =   2700
   ClientLeft      =   2415
   ClientTop       =   2205
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Height          =   3105
   Left            =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4575
   Top             =   1860
   Width           =   4695
   Begin VB.Timer Timer3 
      Interval        =   30
      Left            =   3840
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   120
      Left            =   2040
      Top             =   1560
   End
   Begin VB.Timer Yes 
      Interval        =   100
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   0
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         name            =   "돋움체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "ENDQUES.frx":0000
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image OrImg 
      Height          =   735
      Left            =   1800
      Picture         =   "ENDQUES.frx":0006
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Image NoImg 
      Height          =   615
      Left            =   3120
      Picture         =   "ENDQUES.frx":145C
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Image YesImg 
      Height          =   720
      Left            =   360
      Picture         =   "ENDQUES.frx":2612
      Top             =   1560
      Width           =   1305
   End
End
Attribute VB_Name = "EndQues"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Mode1
Dim Mode2
Dim Mode3

Private Sub NoImg_Click()
EndQues.Hide
MainFrm.Show
End Sub

Private Sub OrImg_Click()
Beep
Beep
msg = "Or을 눌렀음!" + Chr(13) + Chr(10) + "컴퓨터 : 왜 Or 눌렀냐?" + Chr(13) + Chr(10) + "나 : 야! 날 무로 보지마! "
I = vbInformation
Title = "~~~~ Or누름 귀신 바이러스~~~ "
Res = MsgBox(msg, I, Title)

End Sub

Private Sub Timer1_Timer()
'랜덤 계시판 메세지 Timer
Randomize
A = Int(Rnd * 12)
If A = 1 Then
  Tmp = "숫자야구! 인택이가 만듭니다! E-Mail : dark_hani@yahoo.co.kr 과 한메일 그리고 인티즌 까지..."
ElseIf A = 2 Then
   Tmp = "만든이 E-mail : dark_hani@hanmail.net 와 인티즌에도..."
ElseIf A = 3 Then
   Tmp = "인천 광성 고등학교 Hacker System 8기 한인택 입니다!"
ElseIf A = 4 Then
   Tmp = "좀 더 열심히 제작에 임하겠습니다."
ElseIf A = 5 Then
   Tmp = " 끝없는 배움터! 정열의 가르침을 주시는 교사진! 인천 광성고등학교!"
ElseIf A = 6 Then
   Tmp = " 명문대학진학 100% 신화! 광성 고등학교가 만들어 갑니다!"
ElseIf A = 7 Then
   Tmp = "좀 더 멋지고 신나는 컴퓨팅 환경을 위하여~ Hacker는 멈추지 않습니다!"
ElseIf A = 8 Then
   Tmp = "만든이 정보 : Home Page = avenger.inticity.com 과 E-Mail은 avenger@intizen.com"
ElseIf A = 9 Then
   Tmp = "여러분의 편의를 생각하여 인티즌, 한메일, 야후 이렇게 세곳에 연락 E-mail이 있습니다. 편하신 곳으로 연락 주세요!"
 Else
   Tmp = "한메일 : dark_hani@hanmail.net , 야후 : dark_hani@yahoo.co.kr"
End If
Text1.Text = Tmp
End Sub

Private Sub Timer2_Timer()
If Mode2 = 1 Then
  OrImg.Top = OrImg.Top - 20
  If OrImg.Top < 1560 Then
    Mode2 = 2
  End If
Else
  OrImg.Top = OrImg.Top + 20
  If OrImg.Top > 2040 Then
    Mode2 = 1
  End If
End If

End Sub

Private Sub Timer3_Timer()
If Mode3 = 1 Then
  NoImg.Top = NoImg.Top + 20
  If NoImg.Top >= 2040 Then
    Mode3 = 2
  End If
Else
  NoImg.Top = NoImg.Top - 20
  If NoImg.Top <= 1560 Then
    Mode3 = 1
  End If
End If

End Sub

Private Sub Yes_Timer()
If Mode1 = 1 Then
  YesImg.Top = YesImg.Top + 20
  If YesImg.Top >= 2040 Then
    Mode1 = 2
  End If
Else
  YesImg.Top = YesImg.Top - 20
  If YesImg.Top <= 1560 Then
    Mode1 = 1
  End If
End If

End Sub

Private Sub YesImg_Click()
Rem shell="c:\windows\rundll.exe user.exe,exitwindows"
End
End Sub


