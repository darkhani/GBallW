VERSION 4.00
Begin VB.Form SiljunFrm 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "���ǽ�������"
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
      Caption         =   "��  ��"
      BeginProperty Font 
         name            =   "����"
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
         name            =   "����ü"
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
      Caption         =   "��   ��"
      BeginProperty Font 
         name            =   "����ü"
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
      Caption         =   "��  ��"
      BeginProperty Font 
         name            =   "����ü"
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
         name            =   "����ü"
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
         name            =   "����ü"
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
         name            =   "����ü"
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
         name            =   "����ü"
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
      Caption         =   "��������"
      BeginProperty Font 
         name            =   "����ü"
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
         name            =   "����ü"
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
Dim A, B, C, D, e  '��
Dim F, G, H, I, J '��
Dim S, Ba          '�Ǵ�
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
   msg = "�װ� ��ȣ ������? ������ �ϼ���!"   ' Define message.
   Style = vbInformation
   Title = "�� ���߼̽��ϴ�!"  ' Define title.
   Response = MsgBox(msg, Style, Title)
Exit Sub
End Sub
Private Sub PanDan()
S = 0: Ba = 0
  If F = A Then ' ��Ʈ����ũ ó�� ��ƾ
    S = S + 1
  End If
  If G = B Then
    S = S + 1
  End If
  If H = C Then
    S = S + 1
  End If
  
  If F = B Then '��
   Ba = Ba + 1
  End If
  If F = C Then
    Ba = Ba + 1
  End If
  
  If G = A Then '��
   Ba = Ba + 1
  End If
  If G = C Then
    Ba = Ba + 1
  End If
  
  If H = A Then '��
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

Private Sub Command12_Click() '�� �Ǵ�
F = Val(FirstNum.Text)
G = Val(SecondNum.Text)
H = Val(ThreeNum.Text)
If Foor = 1 Then
 I = Val(FourNum.Text)
End If
Close #1
If Level = 0 Then  '�� 0�̸�
 If F = 5 And G = 2 And H = 3 Then 'ù�� ��
   
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "1"
   Close #1
   Ne.Visible = True
   Text1.Text = "�����մϴ�! 1�ܰ��ȣ�� ã�ұ���! : " & F & G & H
   Exit Sub
Else
   Fall
 End If
End If '���� 0 �ϰ�� �ݱ�

If Level = 1 Then  '�� 1�̸�
 If F = 4 And G = 1 And H = 9 Then '2�� ��
   Ne.Visible = True
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "2"
   Close #1
   
   Text1.Text = "�����մϴ�! 2�ܰ��ȣ�� ã�ұ���! : " & F & G & H
   Exit Sub
  Else
   Fall
 End If
End If '���� 1 �ϰ�� �ݱ�

If Level = 2 Then '���� 2
   
  A = 6  '��ǻ�Ͱ� �����ϰ� �ִ� �� 637
  B = 3
  C = 7
   
  If F = 6 And G = 3 And H = 7 Then '����
   
   Ne.Visible = True
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "3"
   Close #1
    Text1.Text = "�����մϴ�! 2�ܰ��ȣ�� ã�ұ���! : " & F & G & H
  
  Else ' �ƴϸ�
    S = 0: Ba = 0
    PanDan
    P = Str(F)
    Q = Str(G)
    n = Str(H)
    n2 = Str(S)
    n3 = Str(Ba)
    Tmp = "�Է��Ͻ� " & P & Q & n + "��  " + n2 + " S " + n3 + " B " + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + Tmp

  End If
 End If '����2 �ݱ�

If Level = 3 Then  '�� 3�̸�
 
 If S = 3 Then '3�� ��
   
   Ne.Visible = True
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "4"
   Close #1
   
   Text1.Text = "�����մϴ�! 3�ܰ��ȣ�� ã�ұ���! : " & F & G & H
   Exit Sub
  Else
   
   RanABC
   Label2.Caption = Label2.Caption & A & B & C
   PanDan
   
   If S = 3 Then '3�� ��
   
     Ne.Visible = True
     Open "LEVELW.DAT" For Output As #1 '���� ���� ����
       Print #1, "4"
     Close #1
   
     Text1.Text = "�����մϴ�! 3�ܰ��ȣ�� ã�ұ���! : " & F & G & H
     Exit Sub
   End If

    P = Str(F)
    Q = Str(G)
    n = Str(H)
    n2 = Str(S)
    n3 = Str(Ba)
    Tmp = "���� : " & P & Q & n + "��  " + n2 + " S " + n3 + " B �̴�!!" + Chr(13) + Chr(10)
    Text1.Text = Text1.Text + Tmp

 End If

End If '���� 3 �ϰ�� �ݱ�

If Level = 4 Then  '�� 4�̸�
 If F = 8 And G = 1 And H = 5 Then '4�� ��
   
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "5"
   Close #1
   Ne.Visible = True
   Text1.Text = "�����մϴ�! 4�ܰ��ȣ�� ã�ұ���! : " & F & G & H
   Exit Sub
Else
   Fall
 End If
End If '���� 4 �ϰ�� �ݱ�

If Level = 5 Then  '�� 5�̸�
 If F = 2 And G = 4 And H = 8 Then '5�� ��
   
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "6"
   Close #1
   Ne.Visible = True
   Text1.Text = "�����մϴ�! 5�ܰ��ȣ�� ã�ұ���! : " & F & G & H
   Exit Sub
Else
   Fall
 End If
End If '���� 5 �ϰ�� �ݱ�

If Level = 6 Then  '�� 6�̸�
 If F = 1 And G = 4 And H = 3 Then '6�� ��
   
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "7"
   Close #1
   Ne.Visible = True
   Text1.Text = "�����մϴ�! 6�ܰ��ȣ�� ã�ұ���! : " & F & G & H
    Foor = 1
   Exit Sub
Else
   Fall
 End If
End If '���� 6 �ϰ�� �ݱ�

If Level = 7 Then  '�� 7�̸�
Foor = 1
 If F = 5 And G = 4 And H = 2 And I = 8 Then '7�� ��
   
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "8"
   Close #1
   Ne.Visible = True
   Text1.Text = "�����մϴ�! 7�ܰ��ȣ�� ã�ұ���! : " & F & G & H & I
   Exit Sub
Else
   Fall
 End If
End If '���� 7 �ϰ�� �ݱ�

If Level = 8 Then  '�� 8�̸�
 If F = 9 And G = 6 And H = 5 And I = 1 Then '8�� ��
   
   Open "LEVELW.DAT" For Output As #1 '���� ���� ����
     Print #1, "8" '���� �߰��� �ø���
   Close #1
   Ne.Visible = True
   Text1.Text = "�����մϴ�! 8�ܰ��ȣ�� ã�ұ���! : " & F & G & H & I + "������ �� ������ ������� �Դϴ�."
   Foor = 0
   Exit Sub
Else
   Fall
 End If
End If '���� 8 �ϰ�� �ݱ�
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
Label2.Caption = "����� " + Q$ + "�ܰ�"
 
 If Level = 0 Then
    Text1.Text = "���ھ߱� ������ Dark_hani�� ����ǰ�� ���ھ߱� ������ ģ���鿡�� �غ���� ��Ź�ߴ�. �����߿�" + Chr(13) + Chr(13) + Chr(10) + " 123 ---> 2S , 456 ---> 1B ,789�� 147 OUT!" + Chr(13) + Chr(13) + Chr(10) + " �̰���� �̿��ؼ� �����ܰ�� �Ѿ� ������!"
 
 ElseIf Level = 1 Then
    Text1.Text = "�ָ𾾴� ǲ���� ���α׷��Ӵ�. �ָ𾾴� �����̿��� ���� ���߾� ����鼭 ������ �Ѵ�. ������ ������ ����." + Chr(13) + Chr(13) + Chr(10) + "123 -> 1B , 456 -> 1S , 789 -> 1S" + Chr(13) + Chr(13) + Chr(10) + "����� ������ �巯�� ����� �����Ϸ��� �� �ʿ�� ����! ���� �̸� Ǯ�� �ָ��� �ڸ� ��¦�ϰ� �ϴ� ���̴�." + Chr(13) + Chr(13) + Chr(10) + "Hint - �̰� �����̴�! ó���� 4, �ݺ��Ǵ°� ����. "
 
 ElseIf Level = 2 Then
  msg = "��õ�� �ִ� �� �����б� ������ ��Ŀ ������ 20���� ��Ƴ��� �� �б� ���� ��ǻ�Ϳ� ħ���� ����� ��ȸ�� ������. ������ �� �׷������� ������ �� ��ȸ�� ������ ��� �ĺ��� ����� ��й�ȣ�� �˾Ƴ��� �ȴ�. �׷��� ������ ū�ڸ� ��¦�ϰ�..."
  Style = vbInformation
  Title = "�� 2�ܰ� ����"  ' Define title.
  Response = MsgBox(msg, Style, Title)
  Tmp = "���ھ߱��� Ǯ���� ��ȣ�� ã������!" + Chr(13) + Chr(10)
  Text1.Text = Tmp
  
  ElseIf Level = 3 Then
   Text1.Text = "���� : ���ھ�! ���� ������ ���߸� ���� �ܰ� ��ȣ�� �˷� ����.." + Chr(13) + Chr(10) + "���� : �˰ڽ��ϴ�!" + Chr(13) + Chr(10) + "���� : ���ڴ� 3�ڸ�����!" + Chr(13) + Chr(10)
  
  ElseIf Level = 4 Then
   Text1.Text = "���� ������ �밡�ٰ� �ִ� 4�ܰ� �Դϴ�!" + Chr(13) + Chr(10) + " ���� ���� ���� Ư�� ���񽺷� ���� �帮����!" + Chr(13) + Chr(10) + "�ٵ� ���� ������?" + Chr(13) + Chr(10) + " ����� �ֱ⸦... " + Chr(13) + Chr(10) + " Hint-ù° �ڸ�:¦��, �ι�°:����Ȧ��, ����°:�߰�Ȧ��" + Chr(13) + Chr(13) + Chr(10) + "PS : ������? �ݶ� ���� ���ð� �ϼ���!"
  ElseIf Level = 5 Then '248
   msg = "�̹� �ܰ�� �����б� 1�г� ���� ������ �����μ� ���߰������� �� Ǯ�� �ִ� ���� �Դϴ�!" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "�Ѹ��� ��������� ������ Ǯ�� �μ����� ����� �˾Ƴ´�. �̸� �̿��� ������ �Ѿ� ���ÿ�!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
   msg = msg + "�μ� ���� �� ���� x^2+4x+4�̴�. " + Chr(13) + Chr(10) + "�̽��� ���� ���� �� ���ڸ��� ���̰� �׼��� ���ǻ� X�� ǥ���Ѵ�!" + Chr(13) + Chr(10) + "X 1 4 -> 1S 1B" + Chr(13) + Chr(10) + "8 X 9 -> 2B" + Chr(13) + Chr(10) + "7 5 X -> 1B" + Chr(13) + Chr(10) + "Hint - �μ��� ���� ������ �Ѽ�??"
   Text1.Text = msg
  ElseIf Level = 6 Then '143
   msg = "Ŭ���� ������� ������ ���� �־��� ���̴�." + Chr(13) + Chr(10) + "Ŭ���� : ���������ϴ� ���ڸ� ���� ���Գ�!" + Chr(13) + Chr(10) + "��������� : �׷��� ����!!" + Chr(13) + Chr(10) + "Ŭ���� : ���� 9�� ��� �׸��� 8�� 3�� ���� 2���, 3����� ���� �����Ѵٳ�..." + Chr(13) + Chr(10) + "������ ��� : (��! �߲�..) " + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "������ ��� : �� ���� �ֽÿ�!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
   msg = msg + "Hint - ���� ���� 957�̶�� �����Ǹ� 759��� �Է��ؾ� ���� ó�� �˴ϴ�!" + Chr(13) + Chr(10) + " �� ���� ���� �������� ����...." + Chr(13) + Chr(10)
   Text1.Text = msg
  ElseIf Level = 7 Then '5428
   Foor = 1
   msg = "��ǻ�� ��Ŀ�� ����� �������� ��ŷ �ϴ��� ���� ����� ���¸� ȹ���ߴ�." + Chr(13) + Chr(10) + "��й�ȣ�� ABCD �� ���ڷ� ��ȣȭ �Ǿ� �־���." + Chr(13) + Chr(10) + "�� ���¿� 1���� ��� �־���. ���� �ܼ��� �̿��� �����ܰ� ��ȣ�� ȹ�� �ϼ���!" + Chr(13) + Chr(10)
   msg = msg + " A�� ���� �� ���¿� �� ���׼��� 1/2�� �� ���ڸ� ����" + Chr(13) + Chr(10) + "B�� �� " + Chr(13) + Chr(10) + "C�� [A * (1/A)]�� 1�� ���Ѽ�" + Chr(13) + Chr(10) + "D�� B�� 2�� : 5 < D < 9" + Chr(13) + Chr(10) + "����� �︸ ���ڰ� �Ǳ�..."
   Text1.Text = msg
  ElseIf Level = 8 Then '9651
   Foor = 1
   msg = "ģ���� ��ȭ��ȣ�� �˷� �־���." + Chr(13) + Chr(10) + "�׷��� ����� �� ��ȭ��ȣ�� ���� 4�ڸ��� �ذ� ���Ҵ�." + Chr(13) + Chr(10)
   msg = msg + "�� ������ ���� 270�̰� ���� 21�̴�." + Chr(13) + Chr(10) + "������ �� ���ڸ� ���� 9" + Chr(13) + Chr(10) + "������������ �����ؾ� ���̴�." + Chr(13) + Chr(10) + "���� 1,4,7,9�� ���̶�� �����ϸ� 9741�� �Է��ϸ� ��." + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "����� ģ������ �� ��ȭ�� �� �� �ֱ⸦..."
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
