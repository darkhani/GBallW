VERSION 4.00
Begin VB.Form MainFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   495
   ClientTop       =   810
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Height          =   6735
   Left            =   435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6735
   Top             =   180
   Width           =   6855
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1800
      Top             =   5160
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1200
      Top             =   5160
   End
   Begin VB.Timer Timer2 
      Interval        =   1200
      Left            =   600
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      Caption         =   " 게  시  판  : 해골은 절대 클릭 하지 마세요!"
      BeginProperty Font 
         name            =   "돋움체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   6015
      Begin VB.Label Label4 
         BeginProperty Font 
            name            =   "굴림체"
            charset         =   1
            weight          =   400
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   5160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   6720
      Y1              =   600
      Y2              =   600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   4800
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
   Begin VB.Image HeP 
      Height          =   480
      Left            =   5400
      MouseIcon       =   "GBALL.frx":0000
      MousePointer    =   13  'Custom
      Picture         =   "GBALL.frx":030A
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label KL 
      Caption         =   "Label5"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Image Info 
      Height          =   330
      Left            =   3240
      Picture         =   "GBALL.frx":0614
      Top             =   120
      Width           =   360
   End
   Begin VB.Image VerIn 
      Height          =   330
      Left            =   2880
      Picture         =   "GBALL.frx":079E
      Top             =   120
      Width           =   360
   End
   Begin VB.Image SysInfo 
      Height          =   330
      Left            =   2520
      Picture         =   "GBALL.frx":0928
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Moi 
      Height          =   330
      Left            =   1800
      Picture         =   "GBALL.frx":0AB2
      Top             =   120
      Width           =   360
   End
   Begin VB.Image BallImg 
      Height          =   330
      Left            =   1080
      Picture         =   "GBALL.frx":0C3C
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Door 
      Height          =   330
      Left            =   480
      Picture         =   "GBALL.frx":0DC6
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Jung 
      Height          =   330
      Left            =   120
      Picture         =   "GBALL.frx":0F50
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Ball 
      Height          =   750
      Left            =   6480
      Picture         =   "GBALL.frx":10DA
      Top             =   4800
      Width           =   22500
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3840
      X2              =   3840
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3840
      X2              =   6600
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "돋움체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   5640
      Y2              =   6000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6720
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         name            =   "바탕체"
         charset         =   1
         weight          =   400
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   6735
   End
   Begin VB.Menu MnuSet 
      Caption         =   "환경"
      Begin VB.Menu MnuNum2 
         Caption         =   "숫자야구란"
      End
      Begin VB.Menu MnuHi1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnd 
         Caption         =   "끝          ALT + X"
      End
   End
   Begin VB.Menu MnuNum1 
      Caption         =   "숫자야구"
      Begin VB.Menu MnuSt 
         Caption         =   "게임시작    F5"
      End
   End
   Begin VB.Menu MnuMoi 
      Caption         =   "실전 G-Ball"
      Begin VB.Menu MnuMoi2 
         Caption         =   "모의 실전 게임"
      End
      Begin VB.Menu MnuMoi3 
         Caption         =   "모의 실전 설명"
      End
      Begin VB.Menu Format 
         Caption         =   "모의실전 LEVEL 초기화"
      End
   End
   Begin VB.Menu MnuAc 
      Caption         =   "악세서리"
      Begin VB.Menu MnuSI 
         Caption         =   "시스템 정보"
      End
      Begin VB.Menu MnuHi3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerUp 
         Caption         =   "버젼 업 과정"
      End
      Begin VB.Menu MnuInfo 
         Caption         =   "정보                       F1"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu End9 
         Caption         =   "끝         ESC"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Shared Motion
Dim LStyle
Dim m_danchu      '마우스 단추를 눌렀으면 True
                  '놓았으면 False
Dim MouseMo(2)

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_danchu = True
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_danchu = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If m_danchu = False Then         '단추가 눌려져 있지 않으면
      If y >= 120 And y <= 120 + 330 Then
       If x >= 120 And x <= 120 + 360 Then
             Label1.Caption = "숫자야구 역사?": MouseMo(0) = 1
             Exit Sub
        ElseIf x >= 480 And x <= 480 + 360 Then
             Label1.Caption = "끝냅니다!": MouseMo(0) = 2
             Exit Sub
        ElseIf x >= 1080 And x <= 1080 + 360 Then
             Label1.Caption = "숫자야구를 직접 즐기세요!": MouseMo(0) = 3
             Exit Sub
        ElseIf x >= 1800 And x <= 1800 + 360 Then
             Label1.Caption = "실시간으로 즐기는 서바이벌게임을 즐기세요!": MouseMo(0) = 4
             Exit Sub
        ElseIf x >= 2520 And x <= 2520 + 360 Then
             Label1.Caption = "시스템 사양 출력기.": MouseMo(0) = 6
             Exit Sub
        ElseIf x >= 2880 And x <= 2880 + 360 Then
             Label1.Caption = "버젼업 정보 출력기.": MouseMo(0) = 7
             Exit Sub
        ElseIf x >= 3240 And x <= 3240 + 360 Then
             Label1.Caption = "숫자야구 만든이 정보": MouseMo(0) = 8
             Exit Sub
        End If
      End If
      MouseMo(0) = 0
    Label1.Caption = "마우스로 해당 메뉴를 클릭하여 작업을 실행 하세요!"
    End If
End Sub
Private Sub BallImg_Click()
MainFrm.Hide
BaseFrm.Show
End Sub

Private Sub Door_Click()
EndQues.Show
MainFrm.Hide
End Sub

Private Sub End9_Click()
EndQues.Show
MainFrm.Hide
End Sub

Private Sub Form_Load()
MainFrm.Hide
InfoFrm.Show
MainFrm.Caption = "숫자야구 '2000 Ver 0.99 [320]"
Label1.Caption = "마우스로 해당 메뉴를 클릭하여 작업을 실행 하세요!"
Motion = 3
LStyle = 0
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 KL.Caption = "No"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
KL.Caption = "Yes" & KeyCode
'Alt+X키를 눌렀으면 실행 종료(88:X의 KeyCode)
    If KeyCode = 88 And Shift = 4 Or KeyCode = 27 Then EndQues.Show: MainFrm.Hide
     
    
    'F1 키를 눌렀으면 도움말 MsgBox를 출력
    If KeyCode = 112 Then AboutFrm.Show
    If KeyCode = 116 Then MainFrm.Hide: BaseFrm.Show

    '눌려진 키에 따라 x,y 값을 변경
    Select Case Shift * 100 + KeyCode
      Case 37: x = x - 50             '좌
      Case 38: y = y - 50             '상
      Case 39: x = x + 50             '우
      Case 40: y = y + 50             '하 키
     
     Case Else:  Exit Sub       '여타키를 눌렀으면
    
    End Select
    
End Sub

Private Sub Format_Click()
Kill "levelw.dat"
Open "LEVELW.DAT" For Output As #1
 Print #1, "0"
Close #1
End Sub

Private Sub HeP_Click()
Dim Vi(9)
Vi(0) = "Dark_Avenger": Vi(1) = "Damanagi": Vi(2) = "CIH/Win '95"
Vi(3) = "Anti_Cmos": Vi(4) = "G2.변형": Vi(5) = "NRLG.변형"
Vi(6) = "Michelangelo": Vi(7) = "Scorpion": Vi(8) = "PS-MPC.변형"
Vi(9) = "Y4"
Beep
Beep
msg = "해골을 눌렀음!" + Vi(Int(Rnd * 8)) + " Virus 에 감염!"
I = vbInformation
Title = "V3를 호출 하세요!"
Res = MsgBox(msg, I, Title)
End Sub

Private Sub Info_Click()
AboutFrm.Show
End Sub

Private Sub Jung_Click()
JungboFrm.Show
End Sub

Private Sub MnuEnd_Click()
EndQues.Show
MainFrm.Hide
End Sub

Private Sub MnuInfo_Click()
AboutFrm.Show
End Sub

Private Sub MnuMoi2_Click()
MainFrm.Hide
SiljunFrm.Show
End Sub

Private Sub MnuMoi3_Click()
MainFrm.Hide
HelpFrm.Show
End Sub

Private Sub MnuNum2_Click()
JungboFrm.Show
End Sub

Private Sub MnuSI_Click()
MainFrm.Hide
SysInfoFrm.Show
End Sub

Private Sub MnuSt_Click()
MainFrm.Hide
BaseFrm.Show
End Sub

Private Sub MnuVerUp_Click()
VerUpFrm.Show
End Sub
Private Sub Moi_Click()
Unload MainFrm
SiljunFrm.Show

End Sub
Private Sub SysInfo_Click()
SysInfoFrm.Show
End Sub
Private Sub Timer1_Timer()
'시간출력 Timer
Label3.Caption = Now
End Sub

Private Sub Timer2_Timer()
'랜덤 계시판 메세지 Timer
Randomize
A = Int(Rnd * 12)
If A = 1 Then
  Tmp = "숫자야구! 인택이가 만듭니다! E-Mail : dark_hani@yahoo.co.kr 과 한메일 그리고 인티즌 까지..."
ElseIf A = 2 Then
   Tmp = "만든이 E-mail : dark_hani@hanmail.net 와 인티즌에도..."
ElseIf A = 3 Then
   Tmp = "해커 시스템 8기 한인택 입니다!"
ElseIf A = 4 Then
   Tmp = "좀 더 열심히 제작에 임하겠습니다."
ElseIf A = 5 Then
   Tmp = " 끝없는 배움터! 정열의 가르침을 주시는 교사진!"
ElseIf A = 6 Then
   Tmp = " 명문대학진학 100% 신화! 광성 고등학교가 만들어 갑니다!"
ElseIf A = 7 Then
   Tmp = "좀 더 좋은 알고리즘으로 찾아 뵙겠습니다."
ElseIf A = 8 Then
   Tmp = "만든이 정보 : Home Page = avenger.inticity.com 과 E-Mail은 avenger@intizen.com"
ElseIf A = 9 Then
   Tmp = "여러분의 편의를 생각하여 인티즌, 한메일, 야후 이렇게 세곳에 연락 E-mail이 있습니다. 편하신 곳으로 연락 주세요!"
 Else
   Tmp = "겉 모습만 화려한 Software 보다는 알차고 실속있는 Software를 개발해겠습니다!"
End If
Label4.Caption = Tmp
End Sub

Private Sub Timer3_Timer()
'베너 출력
Ball.Left = Ball.Left - 10
If Ball.Left < -17760 Then Ball.Left = 6480

End Sub
Private Sub CtrlMove()
Select Case Motion
    Case 1
        ' Move the graphic left and up by 20 twips using the Move method.
        HeP.Move HeP.Left - 20, HeP.Top - 20
        ' If the graphic reaches the left edge of the form, move it to the right and up.
        If HeP.Left <= 0 Then
            Motion = 2
        ' If the graphic reaches the top edge of the form, move it to the left and down.
        ElseIf HeP.Top <= 600 Then
            Motion = 4
        End If
    Case 2
        ' Move the graphic right and up by 20 twips.
        HeP.Move HeP.Left + 20, HeP.Top - 20
        ' If the graphic reaches the right edge of the form, move it to the left and up.
        ' Routine determines the right edge of the form by subtracting the graphic
        ' width from the form width.
        If HeP.Left >= (MainFrm.Width - HeP.Width) Then
            Motion = 1
        ' If the graphic reaches the top edge of the form, move it to the right and down.
        ElseIf HeP.Top <= 600 Then
            Motion = 3
        End If
    Case 3
        ' Move the graphic right and down by 20 twips.
        HeP.Move HeP.Left + 20, HeP.Top + 20
        ' If the graphic reaches the right edge of the form, move it to the left and down.
        If HeP.Left >= (MainFrm.Width - HeP.Width) Then
            Motion = 4
        ' If the graphic reaches the bottom edge of the form, move it to the right and up.
        ' Routine determines the bottom of the form by subtracting
        ' the graphic height from the form height less 680 twips for the height
        ' of title bar and menu bar.
        ElseIf HeP.Top >= (MainFrm.Height - HeP.Height) - 680 Then
            Motion = 2
        End If
    Case 4
        ' Move the graphic left and down by 20 twips.
        HeP.Move HeP.Left - 20, HeP.Top + 20
        ' If the graphic reaches the left edge of the form, move it to the right and down.
        If HeP.Left <= 0 Then
            Motion = 3
        ' If the graphic reaches the bottom edge of the form, move it to the left and up.
        ElseIf HeP.Top >= (MainFrm.Height - HeP.Height) - 680 Then
            Motion = 1
        End If
    End Select

End Sub
Private Sub Timer4_Timer()
 CtrlMove
End Sub

Private Sub VerIn_Click()
VerUpFrm.Show
End Sub




