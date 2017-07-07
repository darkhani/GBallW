VERSION 4.00
Begin VB.Form JungboFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   3  'Fixed Double
   Caption         =   "숫자야구 정보"
   ClientHeight    =   4950
   ClientLeft      =   825
   ClientTop       =   1710
   ClientWidth     =   6615
   Height          =   5355
   Left            =   765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6615
   Top             =   1365
   Width           =   6735
   Begin VB.CommandButton Command1 
      Caption         =   "알았음"
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         name            =   "돋움체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3480
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
End
Attribute VB_Name = "JungboFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
JungboFrm.Hide
MainFrm.Show
End Sub

Private Sub Form_Load()
Open "Baseball.Num" For Input As #1
 Do While Not EOF(1)
    Line Input #1, Fline
    Fline = Fline & Chr(13) & Chr(10)
    Fdata = Fdata & Fline
 Loop
Close #1
Text1.Text = Fdata
End Sub
