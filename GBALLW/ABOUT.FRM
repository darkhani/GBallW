VERSION 4.00
Begin VB.Form AboutFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   3  'Fixed Double
   Caption         =   "숫자야구 정보"
   ClientHeight    =   4515
   ClientLeft      =   1365
   ClientTop       =   1365
   ClientWidth     =   5190
   Height          =   4920
   Left            =   1305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   5190
   Top             =   1020
   Width           =   5310
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
      Height          =   855
      Left            =   600
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "ABOUT.frx":0000
      Top             =   2520
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O                   K"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "ABOUT.frx":0006
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   360
      X2              =   4800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "바탕체"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "AboutFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
AboutFrm.Hide
End Sub

Private Sub Form_Load()
Label1.Caption = "숫자야구 '2000 Version 1.00" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
Label2.Caption = "(C)Copyright 1999~2000 HACKER SYSTEM co.ltd" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "메인 프로그래밍 제작 : 한인택" + Chr(13) + Chr(10) + "메인 루틴: 한상현, 양대연, 조정만" + Chr(13) + Chr(10) + " 디자인 : 이홍균, 홍인식, 김석교" + Chr(13) + Chr(10) + "테스트 : 권재민, 박현규 " + Chr(13) + Chr(10)
Text1.Text = "HELP : 로고나 프로그램에 쓰이는 그림을 멋지게 그려 주실분은 제 메일로 파일좀 날려 주세요! 보답하겠습니다!"
End Sub

Private Sub Timer1_Timer()
End Sub
