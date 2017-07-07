VERSION 4.00
Begin VB.Form SysInfoFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   3  'Fixed Double
   Caption         =   "시스템 정보"
   ClientHeight    =   3720
   ClientLeft      =   1260
   ClientTop       =   1440
   ClientWidth     =   7125
   ControlBox      =   0   'False
   Height          =   4125
   Left            =   1200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7125
   Top             =   1095
   Width           =   7245
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         name            =   "돋움체"
         charset         =   1
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "시스템 정보"
      BeginProperty Font 
         name            =   "돋움체"
         charset         =   1
         weight          =   700
         size            =   20.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "SysInfoFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
SysInfoFrm.Hide
MainFrm.Show
End Sub

Private Sub Form_Load()

WinFlags = GetWinFlags()
Tmp = Chr(13) + Chr(10)
    
 Ver = WindowsVersion()
  D = DosVersion()
  If D = "7.00" Then
     w = "Windows 95 : "
   Else
     w = "Windows : "
  End If
Tmp = Tmp + w + Ver + "(도스 : " & DosVersion() + ")"
        
        If WinFlags And WF_ENHANCED Then
            Tmp = Tmp + "(Enhanced 모드)" + Chr(13) + Chr(10)
        Else
            Tmp = Tmp + "(Standard 모드)" + Chr(13) + Chr(10)
        End If

If WinFlags And WF_CPU586 Then
   Computer = "586"
ElseIf WinFlags And WF_CPU486 Then
   Computer = "486"
ElseIf WinFlags And WF_CPU386 Then
   Computer = "386"
End If
If WinFlags And WF_80x87 Then
            Computer = Computer & "(Math Coprocessor 장착)" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
End If

        Tmp = Tmp + "컴퓨터 CPU : " + Computer + Chr(13) + Chr(10)
   
' Video info.
        Tmp = Tmp + "그래픽 카드: " & GetSysIni("boot.description", "display.drv") + Chr(13) + Chr(10)
        Tmp = Tmp + " 해 상 도  : " & Screen.Width \ Screen.TwipsPerPixelX & " x " & Screen.Height \ Screen.TwipsPerPixelY & Chr(13) + Chr(10)
        Tmp = Tmp + " 컬    러  : " & DeviceColors((hDC)) & "컬러" + Chr(13) + Chr(10) + Chr(13) + Chr(10)

     If GetSystemMetrics(SM_MOUSEPRESENT) Then
        Tmp = Tmp + " 마 우 스  : " & GetSysIni("boot.description", "mouse.drv") + Chr(13) + Chr(13) + Chr(10)
      Else
        Tmp = Tmp + "마우스를 찾지 못함" + Chr(13) + Chr(13) + Chr(10)
     End If
        Tmp = Tmp + " 네트워크  : " & GetSysIni("boot.description", "network.drv") + Chr(13) + Chr(10)
        Tmp = Tmp + " 언    어  : " & GetSysIni("boot.description", "language.dll") + Chr(13) + Chr(10)
        Tmp = Tmp + " 키 보 드  : " & GetSysIni("boot.description", "keyboard.typ") + Chr(13) + Chr(10)

Text1.Text = Tmp
    End Sub
