VERSION 4.00
Begin VB.Form SiljunFrm 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "모의실전게임"
   ClientHeight    =   5250
   ClientLeft      =   570
   ClientTop       =   855
   ClientWidth     =   8355
   ControlBox      =   0   'False
   Height          =   5655
   Left            =   510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8355
   Top             =   510
   Width           =   8475
   Begin VB.Timer Cf 
      Interval        =   2
      Left            =   1560
      Top             =   4800
   End
   Begin VB.CommandButton Be 
      Caption         =   "이  전"
      BeginProperty Font 
         name            =   "굴림"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox FourNum 
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   700
         size            =   20.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Ne 
      Caption         =   "다   음"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "입  력"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox ThreeNum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   700
         size            =   20.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox SecondNum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   700
         size            =   20.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox FirstNum 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   700
         size            =   20.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   360
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "작전종료"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   150
   End
   Begin VB.Image CfImg 
      Height          =   750
      Left            =   7080
      Picture         =   "MOI.frx":0000
      Top             =   4560
      Width           =   30000
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   -600
      Picture         =   "MOI.frx":18AE2
      Top             =   0
      Width           =   9000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   4800
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
End
Attribute VB_Name = "SiljunFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Level
Dim A, B, C, D, e  '컴
Dim F, G, H, I, J '나
Dim S, Ba          '판단
Dim Foor
Dim RanMode

Private Sub Be_Click()
If Not Level = 0 Then
  Level = Level - 1
  Open "LEVELW.DAT" For Output As #1
 Print #1, Level
Close #1
If Level >= 7 Then FourNum.Visible = True: Foor = 1
If Level < 7 Then FourNum.Visible = False: Foor = 0

End If
If Level = 7 Or Level = 8 Then
 Foor = 1
 Else
 Foor = 0
End If
fa$ = Str(Level)
Open "LEVELW.DAT" For Input As #1
 Line Input #1, fa$
Close #1

FirstNum.Text = ""
SecondNum.Text = ""
ThreeNum.Text = ""
FourNum.Text = ""

Cap
End Sub

Private Sub Cf_Timer()
CfImg.Left = CfImg.Left - 20
If CfImg.Left < -27350 Then CfImg.Left = 7920
End Sub


Private Sub Fall()
   msg = "그게 암호 같나요? 열심히 하세요!"   ' Define message.
   Style = vbInformation
   Title = "못 맞추셨습니다!"  ' Define title.
   Response = MsgBox(msg, Style, Title)
Exit Sub
End Sub
Private Sub PanDan()
S = 0: Ba = 0
  If F = A Then ' 스트라이크 처리 루틴
    S = S + 1
  End If
  If G = B Then
    S = S + 1
  End If
  If H = C Then
    S = S + 1
  End If
  
  If F = B Then '볼
   Ba = Ba + 1
  End If
  If F = C Then
    Ba = Ba + 1
  End If
  
  If G = A Then '볼
   Ba = Ba + 1
  End If
  If G = C Then
    Ba = Ba + 1
  End If
  
  If H = A Then '볼
   Ba = Ba + 1
  End If
  If H = B Then
    Ba = Ba + 1
  End If

End Sub
Private Sub RanABC()
If RanMode = 0 Then
 Randomize
da:
    A = Int(Rnd * 8) + 1
    B = Int(Rnd * 8) + 1
    C = Int(Rnd * 8) + 1
  If A = B And A = C And B = C Then
   GoTo da
  End If
 RanMode = 1
Else
  Exit Sub
End If
End Sub

Private Sub Command12_Click() '답 판단
F = Val(FirstNum.Text)
G = Val(SecondNum.Text)
H = Val(ThreeNum.Text)
If Foor = 1 Then
 I = Val(FourNum.Text)
End If
Close #1
If Level = 0 Then  '렙 0이면
 If F = 5 And G = 2 And H = 3 Then '첫판 답
   
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "1"
   Close #1
   Ne.Visible = True
   Text1.Text = "축하합니다! 1단계압호를 찾았군요! : " & F & G & H
   Exit Sub
Else
   Fall
 End If
End If '레벨 0 일경우 닫기

If Level = 1 Then  '렙 1이면
 If F = 4 And G = 1 And H = 9 Then '2판 답
   Ne.Visible = True
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "2"
   Close #1
   
   Text1.Text = "축하합니다! 2단계압호를 찾았군요! : " & F & G & H
   Exit Sub
  Else
   Fall
 End If
End If '레벨 1 일경우 닫기

If Level = 2 Then '레벨 2
   
  A = 6  '컴퓨터가 생각하고 있는 수 637
  B = 3
  C = 7
   
  If F = 6 And G = 3 And H = 7 Then '정답
   
   Ne.Visible = True
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "3"
   Close #1
    Text1.Text = "축하합니다! 2단계압호를 찾았군요! : " & F & G & H
  
  Else ' 아니면
    S = 0: Ba = 0
    PanDan
    P = Str(F)
    Q = Str(G)
    n = Str(H)
    n2 = Str(S)
    n3 = Str(Ba)
    Tmp = "입력하신 " & P & Q & n + "는  " + n2 + " S " + n3 + " B " + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + Tmp

  End If
 End If '레벨2 닫기

If Level = 3 Then  '렙 3이면
 
 If S = 3 Then '3판 답
   
   Ne.Visible = True
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "4"
   Close #1
   
   Text1.Text = "축하합니다! 3단계압호를 찾았군요! : " & F & G & H
   Exit Sub
  Else
   
   RanABC
   Label2.Caption = Label2.Caption & A & B & C
   PanDan
   
   If S = 3 Then '3판 답
   
     Ne.Visible = True
     Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
       Print #1, "4"
     Close #1
   
     Text1.Text = "축하합니다! 3단계압호를 찾았군요! : " & F & G & H
     Exit Sub
   End If

    P = Str(F)
    Q = Str(G)
    n = Str(H)
    n2 = Str(S)
    n3 = Str(Ba)
    Tmp = "스승 : " & P & Q & n + "는  " + n2 + " S " + n3 + " B 이다!!" + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + Tmp

 End If

End If '레벨 3 일경우 닫기

If Level = 4 Then  '렙 4이면
 If F = 8 And G = 1 And H = 5 Then '4판 답
   
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "5"
   Close #1
   Ne.Visible = True
   Text1.Text = "축하합니다! 4단계압호를 찾았군요! : " & F & G & H
   Exit Sub
Else
   Fall
 End If
End If '레벨 4 일경우 닫기

If Level = 5 Then  '렙 5이면
 If F = 2 And G = 4 And H = 8 Then '5판 답
   
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "6"
   Close #1
   Ne.Visible = True
   Text1.Text = "축하합니다! 5단계압호를 찾았군요! : " & F & G & H
   Exit Sub
Else
   Fall
 End If
End If '레벨 5 일경우 닫기

If Level = 6 Then  '렙 6이면
 If F = 1 And G = 4 And H = 3 Then '6판 답
   
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "7"
   Close #1
   Ne.Visible = True
   Text1.Text = "축하합니다! 6단계압호를 찾았군요! : " & F & G & H
    Foor = 1
   Exit Sub
Else
   Fall
 End If
End If '레벨 6 일경우 닫기

If Level = 7 Then  '렙 7이면
Foor = 1
 If F = 5 And G = 4 And H = 2 And I = 8 Then '7판 답
   
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "8"
   Close #1
   Ne.Visible = True
   Text1.Text = "축하합니다! 7단계압호를 찾았군요! : " & F & G & H & I
   Exit Sub
Else
   Fall
 End If
End If '레벨 7 일경우 닫기

If Level = 8 Then  '렙 8이면
 If F = 9 And G = 6 And H = 5 And I = 1 Then '8판 답
   
   Open "LEVELW.DAT" For Output As #1 '레벨 정보 저장
     Print #1, "8" '레벨 추가시 올리자
   Close #1
   Ne.Visible = True
   Text1.Text = "축하합니다! 8단계압호를 찾았군요! : " & F & G & H & I + "하지만 이 버젼은 여기까지 입니다."
   Foor = 0
   Exit Sub
Else
   Fall
 End If
End If '레벨 8 일경우 닫기
End Sub
Private Sub Command6_Click()
Open "LEVELW.DAT" For Output As #1
 Print #1, Level
Close #1
Unload SiljunFrm
MainFrm.Show
End Sub

Private Sub Cap()
Q$ = Str(Level)
Label2.Caption = "당신은 " + Q$ + "단계"
 
 If Level = 0 Then
    Text1.Text = "숫자야구 개발자 Dark_hani는 자작품인 숫자야구 게임을 친구들에게 해보라고 부탁했다. 실행중에" + Chr(13) + Chr(13) + Chr(10) + " 123 ---> 2S , 456 ---> 1B ,789와 147 OUT!" + Chr(13) + Chr(13) + Chr(10) + " 이결과를 이용해서 다음단계로 넘어 가세요!"
 
 ElseIf Level = 1 Then
    Text1.Text = "최모씨는 풋내기 프로그래머다. 최모씨는 인택이에게 답을 맞추어 보라면서 조롱을 한다. 문제는 다음과 같다." + Chr(13) + Chr(13) + Chr(10) + "123 -> 1B , 456 -> 1S , 789 -> 1S" + Chr(13) + Chr(13) + Chr(10) + "당신은 인택의 드러운 기분을 이해하려고 할 필요는 없다! 단지 이를 풀어 최모씨의 코를 납짝하게 하는 것이다." + Chr(13) + Chr(13) + Chr(10) + "Hint - 이건 혁명이다! 처음수 4, 반복되는거 없음. "
 
 ElseIf Level = 2 Then
  msg = "인천에 있는 한 고등학교 교장은 해커 지망생 20명을 모아놓고 그 학교 서버 컴퓨터에 침입해 보라고 대회를 열었다. 교장이 왜 그랬는지는 모르지만 그 대회에 강력한 우승 후보인 당신은 비밀번호만 알아내면 된다. 그래서 교장의 큰코를 납짝하게..."
  Style = vbInformation
  Title = "제 2단계 문제"  ' Define title.
  Response = MsgBox(msg, Style, Title)
  Tmp = "숫자야구를 풀듯이 암호를 찾으세요!" + Chr(13) + Chr(10)
  Text1.Text = Tmp
  
  ElseIf Level = 3 Then
   Text1.Text = "스승 : 제자야! 나의 문제를 맞추면 다음 단계 암호를 알려 주지.." + Chr(13) + Chr(10) + "제자 : 알겠습니다!" + Chr(13) + Chr(10) + "스승 : 숫자는 3자리란다!" + Chr(13) + Chr(10)
  
  ElseIf Level = 4 Then
   Text1.Text = "여긴 죽음의 노가다가 있는 4단계 입니다!" + Chr(13) + Chr(10) + " 여기 까지 오신 특별 서비스로 답을 드리지요!" + Chr(13) + Chr(10) + "근데 답이 뭐더라?" + Chr(13) + Chr(10) + " 행운이 있기를... " + Chr(13) + Chr(10) + " Hint-첫째 자리:짝수, 두번째:작은홀수, 세번째:중간홀수" + Chr(13) + Chr(13) + Chr(10) + "PS : 열받죠? 콜라 한잔 마시고 하세요!"
  ElseIf Level = 5 Then '248
   msg = "이번 단계는 고등학교 1학년 수학 수준의 문제로서 초중고교생이 다 풀수 있는 문제 입니다!" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "한모군은 공통수학의 정석을 풀다 인수분해 방법을 알아냈다. 이를 이용해 담으로 넘어 가시오!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
   msg = msg + "인수 분해 할 식은 x^2+4x+4이다. " + Chr(13) + Chr(10) + "이식의 양의 근이 맨 앞자리의 수이고 그수를 편의상 X로 표기한다!" + Chr(13) + Chr(10) + "X 1 4 -> 1S 1B" + Chr(13) + Chr(10) + "8 X 9 -> 2B" + Chr(13) + Chr(10) + "7 5 X -> 1B" + Chr(13) + Chr(10) + "Hint - 두수의 곱은 나머지 한수??"
   Text1.Text = msg
  ElseIf Level = 6 Then '143
   msg = "클링턴 대통령이 방한한 날에 있었던 일이다." + Chr(13) + Chr(10) + "클링턴 : 내가좋아하는 숫자를 맞춰 보게나!" + Chr(13) + Chr(10) + "교육부장관 : 그렇게 하죠!!" + Chr(13) + Chr(10) + "클링턴 : 나는 9의 약수 그리고 8과 3을 각각 2등분, 3등분한 수를 좋아한다네..." + Chr(13) + Chr(10) + "교육부 장관 : (앗! 뜨끔..) " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "교육부 장관 : 좀 도와 주시오!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
   msg = msg + "Hint - 만약 답이 957이라고 생각되면 759라고 입력해야 정답 처리 됩니다!" + Chr(13) + Chr(10) + " 즉 구한 수에 역순으로 대입...." + Chr(13) + Chr(10)
   Text1.Text = msg
  ElseIf Level = 7 Then '5428
   Foor = 1
   msg = "컴퓨터 해커인 당신은 모은행을 해킹 하던중 부자 김모씨의 계좌를 획득했다." + Chr(13) + Chr(10) + "비밀번호는 ABCD 네 문자로 암호화 되어 있었다." + Chr(13) + Chr(10) + "그 계좌엔 1억이 들어 있었다. 다음 단서를 이용해 다음단계 암호를 획득 하세요!" + Chr(13) + Chr(10)
   msg = msg + " A는 지금 이 계좌에 든 돈액수의 1/2의 맨 앞자리 숫자" + Chr(13) + Chr(10) + "B는 모름 " + Chr(13) + Chr(10) + "C는 [A * (1/A)]에 1을 더한수" + Chr(13) + Chr(10) + "D는 B의 2배 : 5 < D < 9" + Chr(13) + Chr(10) + "당신이 억만 장자가 되길..."
   Text1.Text = msg
  ElseIf Level = 8 Then '9651
   Foor = 1
   msg = "친구가 전화번호를 알려 주었다." + Chr(13) + Chr(10) + "그러나 당신은 그 전화번호의 뒤의 4자리를 잊고 말았다." + Chr(13) + Chr(10)
   msg = msg + "그 수들의 곱은 270이고 합은 21이다." + Chr(13) + Chr(10) + "참고로 맨 앞자리 수는 9" + Chr(13) + Chr(10) + "내림차순으로 정리해야 답이다." + Chr(13) + Chr(10) + "만약 1,4,7,9가 답이라고 생각하면 9741로 입력하면 됨." + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "당신이 친구에게 꼭 전화를 걸 수 있기를..."
   Text1.Text = msg
 End If
End Sub
Private Sub Form_Load()

Ne.Visible = False
Be.Visible = True
CfImg.Left = 7620
 
 Open "Levelw.dat" For Input As #1
 Line Input #1, fileline
   Level = fileline
 Close #1
If Foor = 0 Or Level < 7 Then
 FourNum.Visible = False
Else
 FourNum = True
End If

RanMode = 0
 Cap
End Sub

Private Sub Image1_Click()
HelpFrm.Show
End Sub

Private Sub Ne_Click()
Level = Level + 1
If Level >= 7 Then FourNum.Visible = True: Foor = 1
If Level < 7 Then FourNum.Visible = False: Foor = 0
Open "LEVELW.DAT" For Output As #1
 Print #1, Level
Close #1

Ne.Visible = False
FirstNum.Text = ""
SecondNum.Text = ""
ThreeNum.Text = ""
FourNum.Text = ""
Cap
End Sub

