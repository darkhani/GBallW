VERSION 4.00
Begin VB.Form VerUpFrm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "버젼 업 정보"
   ClientHeight    =   5415
   ClientLeft      =   1395
   ClientTop       =   1350
   ClientWidth     =   6315
   ControlBox      =   0   'False
   Height          =   5820
   Left            =   1335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6315
   Top             =   1005
   Width           =   6435
   Begin VB.CommandButton Command1 
      Caption         =   "확   인"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         name            =   "굴림체"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   240
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "VERUPFRM.frx":0000
      Top             =   240
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   5280
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
   End
End
Attribute VB_Name = "VerUpFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
VerUpFrm.Hide
MainFrm.Show
End Sub

Private Sub Form_Load()
Open "GBALLW.VER" For Input As #1
 Do While Not EOF(1)
    Line Input #1, Fline
    Fline = Fline & Chr(13) & Chr(10)
    Fdata = Fdata & Fline
 Loop
Close #1
Text1.Text = Fdata
End Sub
