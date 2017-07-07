VERSION 4.00
Begin VB.Form AboutFrm 
   Auto3D          =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "코믹시계 버젼 0.9"
   ClientHeight    =   1965
   ClientLeft      =   285
   ClientTop       =   780
   ClientWidth     =   2715
   Height          =   2370
   Left            =   225
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   2715
   Top             =   435
   Width           =   2835
   Begin VB.CommandButton Command1 
      Caption         =   "O K !"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         name            =   "VoyLCD"
         charset         =   1
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "AboutFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
AboutFrm.Hide
End Sub

Private Sub Form_Load()
Text1.Text = "만든이 : 한인택  ****************  도움말             왼쪽눈 : 시간        오른쪽눈 : 분        코 : 초                입 : 디지털 시간"
End Sub
