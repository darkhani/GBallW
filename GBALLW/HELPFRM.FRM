VERSION 4.00
Begin VB.Form HelpFrm 
   Auto3D          =   -1  'True
   Caption         =   "모의실전 설명"
   ClientHeight    =   3615
   ClientLeft      =   720
   ClientTop       =   1725
   ClientWidth     =   7275
   ControlBox      =   0   'False
   Height          =   4020
   Left            =   660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7275
   Top             =   1380
   Width           =   7395
   Begin VB.CommandButton Command11 
      Caption         =   "돌아가기"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   3120
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
End
Attribute VB_Name = "HelpFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Level

Private Sub Command11_Click()
SiljunFrm.Show
Unload HelpFrm
End Sub

Private Sub Cap()
 If Level = 0 Then
  Text1.Text = "123에서 2S 이니까 123중 2개가 맞죠.. 글구 OUT이란 그 숫자들은 쓰이지 않았다는 것을 의미 합니다! 456 이 1B이므로 456중 하나도 답에 포함됩니다."
 End If
 If Level = 1 Then
  Text1.Text = "123이 1B, 456은 1S , 789 는 1S 라고 한다. 그리고 혁명??? 맨 앞자리는 4 , 1960년에 이 혁명이 있었음... "
 End If
 If Level = 2 Then
  Text1.Text = "이번 단계는 해킹 대회라고 하죠~~ 하지만 직접 3개 짜리 숫자야구를 풀어야 하는 단계 입니다. 풀면 그게 답이죠~~ 맨 앞자리 수가 6이던가??? 한번 해보세요!  행운을 빕니다."
 End If
 If Level = 3 Then
  Text1.Text = "스승님의 문제를 맞춰야 한다! 그래야 다음 레벨로 넘어간다? 하! 이건 너무 한다. 정해진 답이 없다니... 답을 찾으세요! 호호호~~~"
 End If
If Level = 4 Then
  Text1.Text = "이번엔 서비스 판으로 그냥 답을 주는 단계 이군요~~ 콜라나 마시고 하라구??? 하지만 만든이는 절대 필요없는 말은 써놓지 않는 답니다. 힌트 - 아래한글과도 연관이... 맨 앞자리가 7보다 크다더군..."
 End If
 If Level = 5 Then
  Text1.Text = "만약 중3이상이라면 그냥 인수분해 해요~ 그 아래라면 ... 힌트 - 8이 들어가는것 같네요~~~"
 End If
 If Level = 6 Then
  Text1.Text = "클링턴 대통령이 좋아하는 수를 구하라구? 참내.. 별... 주어진 수 8,3을 나누면 답이 보인다. 그다음에 역순으로 대입만 하면...."
 End If
 If Level = 7 Then
  Text1.Text = "은행에 1억이 있다. 반을 나누면 5천만원의 맨 앞자리수? 5네... 나머지는 알아서..."
 End If
 If Level = 8 Then
  Text1.Text = "맨 앞자리는 9이구... 4개의 수의 곱은 270이고 합은 21입니다."
 End If
End Sub


Private Sub Form_Load()
Open "Levelw.dat" For Input As #1
 Line Input #1, fileline
   Level = fileline
 Close #1
Cap
End Sub
