VERSION 4.00
Begin VB.Form BaseFrm 
   Auto3D          =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ھ߱�"
   ClientHeight    =   6270
   ClientLeft      =   1770
   ClientTop       =   900
   ClientWidth     =   4965
   ControlBox      =   0   'False
   Height          =   6675
   Left            =   1710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   4965
   Top             =   555
   Width           =   5085
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   5640
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   5640
   End
   Begin VB.TextBox FiveNum 
      BeginProperty Font 
         name            =   "�ü�ü"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox FourNum 
      BeginProperty Font 
         name            =   "�ü�ü"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton InCode 
      Caption         =   "�Է�"
      BeginProperty Font 
         name            =   "����ü"
         charset         =   1
         weight          =   700
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox ThirdNum 
      BeginProperty Font 
         name            =   "�ü�ü"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox SecondNum 
      BeginProperty Font 
         name            =   "�ü�ü"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox FirstNum 
      BeginProperty Font 
         name            =   "�ü�ü"
         charset         =   1
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Dis 
      BeginProperty Font 
         name            =   "����ü"
         charset         =   1
         weight          =   400
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   5640
   End
   Begin VB.PictureBox RetPic 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   1800
      Picture         =   "BASEFRM.frx":0000
      ScaleHeight     =   900
      ScaleWidth      =   1350
      TabIndex        =   3
      Top             =   5280
      Width           =   1350
   End
   Begin VB.PictureBox W5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "BASEFRM.frx":19D2
      ScaleHeight     =   1500
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   3480
      Width           =   4530
   End
   Begin VB.PictureBox W4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "BASEFRM.frx":179A4
      ScaleHeight     =   1500
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   1800
      Width           =   4530
   End
   Begin VB.PictureBox W3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "BASEFRM.frx":2D976
      ScaleHeight     =   1500
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   120
      Width           =   4530
   End
End
Attribute VB_Name = "BaseFrm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim A, B, C, D, e '��ǻ�� ����
Dim F, G, H, I, J ' ���� ����

Dim S, Ba        ' ��� ����
Dim Tmp, Tmp2, Tmp3, Tmp4, Tmp5
Dim Mode

Private Sub Form_Load()
Dis.Visible = False
InCode.Visible = False
End Sub

Private Sub InCode_Click()
S = 0: Ba = 0

If Mode = 3 Then

 
 If FirstNum = "" Or SecondNum = "" Or ThirdNum = "" Then
   Dis.Text = " ���� �Է� ���� : ��ĭ���� ���ڸ� ��������!"
   GoTo e
 End If ' ��ĭ ���� ���� �ݱ�

Tmp = FirstNum.Text
 F = Val(Tmp)
Tmp1 = SecondNum.Text
 G = Val(Tmp1)
Tmp2 = ThirdNum.Text
 H = Val(Tmp2)
 
 If A = F Then S = S + 1  ' �Ǵ� ��Ʈ ��
 If A = G Then Ba = Ba + 1
 If A = H Then Ba = Ba + 1
 
 If B = G Then S = S + 1
 If B = F Then Ba = Ba + 1
 If B = H Then Ba = Ba + 1
 
 If C = H Then S = S + 1
 If C = F Then Ba = Ba + 1
 If C = G Then Ba = Ba + 1
 
   
   If S = 3 Then '���߸�?
      Dis.Text = "��� �մϴ�!" + Chr(13) + Chr(10) + "������ �׸��� ������!"
      Mode = 9
       GoTo e
   End If '���߸� ���κ�
 
Tmp = FirstNum.Text + SecondNum.Text + ThirdNum.Text
Tmp1 = Str(S)
Tmp2 = Str(Ba)
Tmp5 = Tmp + Tmp1 + "S" + Tmp2 + "B" + Chr(13) + Chr(10)

Dis.Text = Dis.Text + Tmp5
End If '���3 ���� �ݱ�

'**************************************************

If Mode = 4 Then
 
 If FirstNum = "" Or SecondNum = "" Or ThirdNum = "" Or FourNum = "" Then
   
   Dis.Text = " ���� �Է� ���� : ��ĭ���� ���ڸ� ��������!"
 
 End If ' ��ĭ ���� ���� �ݱ�
 
Tmp = FirstNum.Text
 F = Val(Tmp)
Tmp1 = SecondNum.Text
 G = Val(Tmp1)
Tmp2 = ThirdNum.Text
 H = Val(Tmp2)
Tmp3 = FourNum.Text
 I = Val(Tmp3)
 
 If A = F Then S = S + 1  ' �Ǵ� ��Ʈ ��
 If A = G Then Ba = Ba + 1
 If A = H Then Ba = Ba + 1
 If A = I Then Ba = Ba + 1
 
 If B = G Then S = S + 1
 If B = F Then Ba = Ba + 1
 If B = H Then Ba = Ba + 1
 If B = I Then Ba = Ba + 1
 
 If C = H Then S = S + 1
 If C = F Then Ba = Ba + 1
 If C = G Then Ba = Ba + 1
 If C = I Then Ba = Ba + 1
 
 If D = I Then S = S + 1
 If D = F Then Ba = Ba + 1
 If D = G Then Ba = Ba + 1
 If D = H Then Ba = Ba + 1
   
   If S = 4 Then '���߸�?
      Dis.Text = "��� �մϴ�!" + Chr(13) + Chr(10) + "������ �׸��� ������!"
      Mode = 9
       GoTo e
   End If '���߸� ���κ�

Tmp = FirstNum.Text + SecondNum.Text + ThirdNum.Text + FourNum.Text
Tmp1 = Str(S)
Tmp2 = Str(Ba)
Tmp5 = Tmp + Tmp1 + "S" + Tmp2 + "B" + Chr(13) + Chr(10)

Dis.Text = Dis.Text + Tmp5
End If '���4 ���� �ݱ�
'*************************************************
If Mode = 5 Then
 
 If FirstNum = "" Or SecondNum = "" Or ThirdNum = "" Or FiveNum = "" Then
   
   Dis.Text = " ���� �Է� ���� : ��ĭ���� ���ڸ� ��������!"
 
 End If ' ��ĭ ���� ���� �ݱ�
 
 Tmp = FirstNum.Text
 F = Val(Tmp)
Tmp1 = SecondNum.Text
 G = Val(Tmp1)
Tmp2 = ThirdNum.Text
 H = Val(Tmp2)
Tmp3 = FourNum.Text
 I = Val(Tmp3)
Tmp4 = FiveNum.Text
 J = Val(Tmp4)
  
 If A = F Then S = S + 1  ' �Ǵ� ��Ʈ ��
 If A = G Then Ba = Ba + 1
 If A = H Then Ba = Ba + 1
 If A = I Then Ba = Ba + 1
 If A = J Then Ba = Ba + 1

 If B = G Then S = S + 1
 If B = F Then Ba = Ba + 1
 If B = H Then Ba = Ba + 1
 If B = I Then Ba = Ba + 1
 If B = J Then Ba = Ba + 1
 
 If C = H Then S = S + 1
 If C = F Then Ba = Ba + 1
 If C = G Then Ba = Ba + 1
 If C = I Then Ba = Ba + 1
 If C = J Then Ba = Ba + 1

 
 If D = I Then S = S + 1
 If D = F Then Ba = Ba + 1
 If D = G Then Ba = Ba + 1
 If D = H Then Ba = Ba + 1
 If D = J Then Ba = Ba + 1

 If e = J Then S = S + 1
 If e = F Then Ba = Ba + 1
 If e = G Then Ba = Ba + 1
 If e = H Then Ba = Ba + 1
 If e = I Then Ba = Ba + 1
  
 If S = 5 Then '���߸�?
      Dis.Text = "��� �մϴ�!" + Chr(13) + Chr(10) + "������ �׸��� ������!"
      Mode = 9
    GoTo e
  End If
 
 Tmp = FirstNum.Text + SecondNum.Text + ThirdNum.Text + FourNum.Text + FiveNum.Text
Tmp1 = Str(S)
Tmp2 = Str(Ba)
Tmp5 = Tmp + Tmp1 + "S" + Tmp2 + "B" + Chr(13) + Chr(10)

Dis.Text = Dis.Text + Tmp5

End If '���5 ���� �ݱ�
e:

FirstNum.Text = ""
SecondNum.Text = ""
ThirdNum.Text = ""
FourNum.Text = ""
FiveNum.Text = ""

End Sub

Private Sub retpic_Click()
 Timer3.Interval = 50
 End Sub

Private Sub Timer1_Timer()
RetPic.Left = RetPic.Left - 100
 If RetPic.Left < 0 Then
   RetPic.Left = 0
 End If

If Mode = 3 Then
   W5.Left = W5.Left + 100
   W4.Left = W4.Left - 100
     If W5.Left > 5000 Then
         
         W3.Top = W3.Top + 100
        
        If W3.Top > 3500 Then
           W3.Top = 3500
           Dis.Left = 120
           Dis.Top = 120
           Dis.Height = 3000
           Dis.Width = 4700
           Dis.Visible = True
           FirstNum.Visible = True
           SecondNum.Visible = True
           ThirdNum.Visible = True
           InCode.Visible = True

     Randomize
Ip:
       A = Int(Rnd * 9) + 1
       B = Int(Rnd * 9) + 1
       C = Int(Rnd * 9) + 1
      If A = B Or B = C Or C = A Then GoTo Ip
      
       Timer1.Interval = 0
      
      
End If
End If
End If '��� 3 �ݱ� ����

If Mode = 4 Then '��� 4 ���� ����
   W5.Left = W5.Left + 100
   W3.Left = W3.Left - 100
     If W5.Left > 5000 Then
         
         W4.Top = W4.Top + 100
        
        If W4.Top > 3500 Then
           W4.Top = 3500
           Dis.Left = 120
           Dis.Top = 120
           Dis.Height = 3000
           Dis.Width = 4700
           Dis.Visible = True
           FirstNum.Visible = True
           SecondNum.Visible = True
           ThirdNum.Visible = True
           FourNum.Visible = True
           InCode.Visible = True

     Randomize
Ip2:
       A = Int(Rnd * 9) + 1
       B = Int(Rnd * 9) + 1
       C = Int(Rnd * 9) + 1
       D = Int(Rnd * 9) + 1
     If A = B Or A = C Or A = D Or B = C Or B = D Or C = D Then GoTo Ip2
       
       Timer1.Interval = 0
      
      
End If
End If

End If '���4 �ݱ� ����
If Mode = 5 Then '���5 ���� ����

   W4.Left = W4.Left + 100
   W3.Left = W3.Left - 100
     If W4.Left > 5000 Then
         
           Dis.Left = 120
           Dis.Top = 120
           Dis.Height = 3000
           Dis.Width = 4700
           Dis.Visible = True
           FirstNum.Visible = True
           SecondNum.Visible = True
           ThirdNum.Visible = True
           FourNum.Visible = True
           FiveNum.Visible = True
           InCode.Visible = True

     Randomize
Ip3:
       A = Int(Rnd * 9) + 1
       B = Int(Rnd * 9) + 1
       C = Int(Rnd * 9) + 1
       D = Int(Rnd * 9) + 1
       e = Int(Rnd * 9) + 1
     If A = B Or A = C Or A = D Or A = e Or B = C Or B = D Or B = e Or C = D Or C = e Or D = e Then GoTo Ip3
       
       Timer1.Interval = 0
End If
End If '���5 �ݱ� ����
End Sub

Private Sub Timer2_Timer()
If Mode = 0 Then '�ﳪ�� ���� ���ư���
  Dis.Visible = False
  
  FirstNum.Visible = False
  SecondNum.Visible = False
  ThirdNum.Visible = False
  InCode.Visible = False

   W3.Top = W3.Top - 100
   If W3.Top < 120 Then
     W3.Top = 120
    
    W5.Left = W5.Left - 100
    W4.Left = W4.Left + 100
    If W4.Left > 120 Then
        Unload BaseFrm
        BaseFrm.Show
       Timer2.Interval = 0
     End If
   End If

End If '���0 �ݱ� ����

If Mode = 8 Then '�糪�󿡼� ���ư���
 
  Dis.Visible = False
  FirstNum.Visible = False
  SecondNum.Visible = False
  ThirdNum.Visible = False
  InCode.Visible = False

   W4.Top = W4.Top - 100
   If W4.Top < 1800 Then
     W4.Top = 1800
    
    W5.Left = W5.Left - 100
    W3.Left = W3.Left + 100
    If W3.Left > 120 Then
        Unload BaseFrm
        BaseFrm.Show
       Timer2.Interval = 0
     End If
   End If

End If '��� 8 �ݱ� ����

If Mode = 7 Then '�����󿡼� ���ư���
 
  Dis.Visible = False
  FirstNum.Visible = False
  SecondNum.Visible = False
  ThirdNum.Visible = False
  InCode.Visible = False

    
    W4.Left = W4.Left - 100
    W3.Left = W3.Left + 100
    If W3.Left > 120 Then
        Unload BaseFrm
        BaseFrm.Show
       Timer2.Interval = 0
     End If

End If '��� 7 �ݱ� ����

End Sub

Private Sub Timer3_Timer()
 RetPic.Left = RetPic.Left - 100
  
  If RetPic.Left < 0 Then
    Timer3.Interval = 0
    Unload BaseFrm
    MainFrm.Show
  End If

End Sub

Private Sub W3_Click()
If Mode = 9 Then
     Timer2.Interval = 50
    Mode = 0
    GoTo Er
End If
Timer1.Interval = 50
Mode = 3
Er:
End Sub

Private Sub W4_Click()
If Mode = 9 Then
   Timer2.Interval = 50
    Mode = 8
    GoTo Er
End If

Timer1.Interval = 50
Mode = 4
Er:
End Sub

Private Sub W5_Click()
If Mode = 9 Then
  Timer2.Interval = 50
  Mode = 7
  GoTo Er
End If

Timer1.Interval = 50
Mode = 5
Er:
End Sub
