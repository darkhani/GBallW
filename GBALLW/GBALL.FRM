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
      Caption         =   " ��  ��  ��  : �ذ��� ���� Ŭ�� ���� ������!"
      BeginProperty Font 
         name            =   "����ü"
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
            name            =   "����ü"
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
         name            =   "����ü"
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
         name            =   "����ü"
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
      Caption         =   "ȯ��"
      Begin VB.Menu MnuNum2 
         Caption         =   "���ھ߱���"
      End
      Begin VB.Menu MnuHi1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnd 
         Caption         =   "��          ALT + X"
      End
   End
   Begin VB.Menu MnuNum1 
      Caption         =   "���ھ߱�"
      Begin VB.Menu MnuSt 
         Caption         =   "���ӽ���    F5"
      End
   End
   Begin VB.Menu MnuMoi 
      Caption         =   "���� G-Ball"
      Begin VB.Menu MnuMoi2 
         Caption         =   "���� ���� ����"
      End
      Begin VB.Menu MnuMoi3 
         Caption         =   "���� ���� ����"
      End
      Begin VB.Menu Format 
         Caption         =   "���ǽ��� LEVEL �ʱ�ȭ"
      End
   End
   Begin VB.Menu MnuAc 
      Caption         =   "�Ǽ�����"
      Begin VB.Menu MnuSI 
         Caption         =   "�ý��� ����"
      End
      Begin VB.Menu MnuHi3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerUp 
         Caption         =   "���� �� ����"
      End
      Begin VB.Menu MnuInfo 
         Caption         =   "����                       F1"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu End9 
         Caption         =   "��         ESC"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Shared Motion
Dim LStyle
Dim m_danchu      '���콺 ���߸� �������� True
                  '�������� False
Dim MouseMo(2)

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_danchu = True
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_danchu = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If m_danchu = False Then         '���߰� ������ ���� ������
      If y >= 120 And y <= 120 + 330 Then
       If x >= 120 And x <= 120 + 360 Then
             Label1.Caption = "���ھ߱� ����?": MouseMo(0) = 1
             Exit Sub
        ElseIf x >= 480 And x <= 480 + 360 Then
             Label1.Caption = "�����ϴ�!": MouseMo(0) = 2
             Exit Sub
        ElseIf x >= 1080 And x <= 1080 + 360 Then
             Label1.Caption = "���ھ߱��� ���� ��⼼��!": MouseMo(0) = 3
             Exit Sub
        ElseIf x >= 1800 And x <= 1800 + 360 Then
             Label1.Caption = "�ǽð����� ���� �����̹������� ��⼼��!": MouseMo(0) = 4
             Exit Sub
        ElseIf x >= 2520 And x <= 2520 + 360 Then
             Label1.Caption = "�ý��� ��� ��±�.": MouseMo(0) = 6
             Exit Sub
        ElseIf x >= 2880 And x <= 2880 + 360 Then
             Label1.Caption = "������ ���� ��±�.": MouseMo(0) = 7
             Exit Sub
        ElseIf x >= 3240 And x <= 3240 + 360 Then
             Label1.Caption = "���ھ߱� ������ ����": MouseMo(0) = 8
             Exit Sub
        End If
      End If
      MouseMo(0) = 0
    Label1.Caption = "���콺�� �ش� �޴��� Ŭ���Ͽ� �۾��� ���� �ϼ���!"
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
MainFrm.Caption = "���ھ߱� '2000 Ver 0.99 [320]"
Label1.Caption = "���콺�� �ش� �޴��� Ŭ���Ͽ� �۾��� ���� �ϼ���!"
Motion = 3
LStyle = 0
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 KL.Caption = "No"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
KL.Caption = "Yes" & KeyCode
'Alt+XŰ�� �������� ���� ����(88:X�� KeyCode)
    If KeyCode = 88 And Shift = 4 Or KeyCode = 27 Then EndQues.Show: MainFrm.Hide
     
    
    'F1 Ű�� �������� ���� MsgBox�� ���
    If KeyCode = 112 Then AboutFrm.Show
    If KeyCode = 116 Then MainFrm.Hide: BaseFrm.Show

    '������ Ű�� ���� x,y ���� ����
    Select Case Shift * 100 + KeyCode
      Case 37: x = x - 50             '��
      Case 38: y = y - 50             '��
      Case 39: x = x + 50             '��
      Case 40: y = y + 50             '�� Ű
     
     Case Else:  Exit Sub       '��ŸŰ�� ��������
    
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
Vi(3) = "Anti_Cmos": Vi(4) = "G2.����": Vi(5) = "NRLG.����"
Vi(6) = "Michelangelo": Vi(7) = "Scorpion": Vi(8) = "PS-MPC.����"
Vi(9) = "Y4"
Beep
Beep
msg = "�ذ��� ������!" + Vi(Int(Rnd * 8)) + " Virus �� ����!"
I = vbInformation
Title = "V3�� ȣ�� �ϼ���!"
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
'�ð���� Timer
Label3.Caption = Now
End Sub

Private Sub Timer2_Timer()
'���� ����� �޼��� Timer
Randomize
A = Int(Rnd * 12)
If A = 1 Then
  Tmp = "���ھ߱�! �����̰� ����ϴ�! E-Mail : dark_hani@yahoo.co.kr �� �Ѹ��� �׸��� ��Ƽ�� ����..."
ElseIf A = 2 Then
   Tmp = "������ E-mail : dark_hani@hanmail.net �� ��Ƽ�𿡵�..."
ElseIf A = 3 Then
   Tmp = "��Ŀ �ý��� 8�� ������ �Դϴ�!"
ElseIf A = 4 Then
   Tmp = "�� �� ������ ���ۿ� ���ϰڽ��ϴ�."
ElseIf A = 5 Then
   Tmp = " ������ �����! ������ ����ħ�� �ֽô� ������!"
ElseIf A = 6 Then
   Tmp = " ������������ 100% ��ȭ! ���� �����б��� ����� ���ϴ�!"
ElseIf A = 7 Then
   Tmp = "�� �� ���� �˰��������� ã�� �˰ڽ��ϴ�."
ElseIf A = 8 Then
   Tmp = "������ ���� : Home Page = avenger.inticity.com �� E-Mail�� avenger@intizen.com"
ElseIf A = 9 Then
   Tmp = "�������� ���Ǹ� �����Ͽ� ��Ƽ��, �Ѹ���, ���� �̷��� ������ ���� E-mail�� �ֽ��ϴ�. ���Ͻ� ������ ���� �ּ���!"
 Else
   Tmp = "�� ����� ȭ���� Software ���ٴ� ������ �Ǽ��ִ� Software�� �����ذڽ��ϴ�!"
End If
Label4.Caption = Tmp
End Sub

Private Sub Timer3_Timer()
'���� ���
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



