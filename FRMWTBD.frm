VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FRMWTBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û��ճ�����"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9330
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   0
      Text            =   "0"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7680
      TabIndex        =   38
      Text            =   "0"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ  ��"
      Height          =   495
      Left            =   1680
      TabIndex        =   33
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ  ��"
      Height          =   495
      Left            =   5640
      TabIndex        =   32
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   9015
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "FRMWTBD.frx":0000
         TabIndex        =   40
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   2040
         TabIndex        =   39
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2040
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   2040
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   6240
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   6240
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   6240
         TabIndex        =   22
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FRMWTBD.frx":006A
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FRMWTBD.frx":00D2
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FRMWTBD.frx":013A
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FRMWTBD.frx":01A6
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FRMWTBD.frx":0212
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FRMWTBD.frx":027E
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2760
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FRMWTBD.frx":02E6
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�û���Ϣ��"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "������û�������˶�ȡ�û���Ϣ"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   120
         Width           =   3855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   8775
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "FRMWTBD.frx":0B93
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "FRMWTBD.frx":0C07
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FRMWTBD.frx":0C6F
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FRMWTBD.frx":0CD7
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "FRMWTBD.frx":0D3F
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FRMWTBD.frx":0DA7
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   6240
            OleObjectBlob   =   "FRMWTBD.frx":0E11
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   3600
      OleObjectBlob   =   "FRMWTBD.frx":0E81
      TabIndex        =   17
      Top             =   120
      Width           =   3615
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8760
      OleObjectBlob   =   "FRMWTBD.frx":0EE6
      Top             =   5280
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   375
      Left            =   6480
      OleObjectBlob   =   "FRMWTBD.frx":111A
      TabIndex        =   34
      Top             =   4560
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   6480
      OleObjectBlob   =   "FRMWTBD.frx":1177
      TabIndex        =   35
      Top             =   4200
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "FRMWTBD.frx":11E1
      TabIndex        =   36
      Top             =   4200
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FRMWTBD.frx":1249
      TabIndex        =   37
      Top             =   4200
      Width           =   3375
   End
End
Attribute VB_Name = "FRMWTBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldpass As String * 4
Dim password(1) As Byte
Dim RdType(5) As Byte
Dim Para(19) As Byte
Dim i As Integer
Dim rst As Recordset, rst1 As Recordset
Dim GDnum As String          '��ǰ�������

Private Sub Command1_Click()
On Error GoTo errhandle
'д��
'���ݲ���
Dim TempGD As String        '������
Dim TempGZ As String        '���ش���
Dim TempTZ As String        '͸֧��
Dim TempShu As String

For i = 0 To 4
    Para(i) = &HFF
Next i
'������
TempGD = FormatString(Val(Text1), 4)
Call ToBCD(Left(TempGD, 2), Para(5))
Call ToBCD(Right(TempGD, 2), Para(6))
'���ش���
TempGZ = FormatString(Val(Text3), 2)
Call ToBCD(TempGZ, Para(7))
'͸֧��
TempTZ = FormatString(Val(Text4), 2)
Call ToBCD(TempTZ, Para(8))
For i = 9 To 12
    Para(i) = &HFF
Next i
TempShu = FormatString(GDnum, 4)
Call ToBCD(Left(TempShu, 2), Para(13))
Call ToBCD(Right(TempShu, 2), Para(14))
For i = 15 To 19
    Para(i) = &HFF
Next i
'�忨1��1-19�ֽ�
st = ser_102(icdev, 1, 0, 22)
If st < 0 Then
    MsgBox ("����ʧ�ܣ�")
    Exit Sub
End If
'д��
st = swr_102_hex(icdev, 1, 1, 20, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'1��������λ����
'*************������λ��0,�˶�����ǰ���ܶ�Ӧ����1���ж�����*****
st = clrrd_102(icdev, 1)
If st < 0 Then
  MsgBox ("������λ�������")
  Exit Sub
End If

'���湺�����ݵ����ݿ�
Dim BUYdate As String   '��ˮ����
  BUYdate = Format(CDate(Now), "yyyy-MM-dd HH:mm:ss")
Dim BUYid As String     '��ˮ���

Set rst = mconn.Execute("select count(yb_id) from wtbddb")
If rst.Fields(0) = 0 Then
    BUYid = "0000001"
Else
    Set rst1 = mconn.Execute("select max(yb_buyid) from WTBDdb")
        If rst1.BOF Then
        Else
        BUYid = FormatString((Val(rst1.Fields(0)) + 1), 7)
        End If
    rst1.Close
End If
rst.Close
Dim BUYnum As String
  BUYnum = FormatString(Str(GDnum), 6)
  
  mconn.Execute ("insert into WTBDdb(yb_buyid,yb_id,yb_type,yb_dn,yb_tdn,yb_num,yb_money,yb_oper,yb_date) values ('" + BUYid + "'," _
                & "'" + Trim(Text15) + "','" + Text7 + "','" + Text1 + "','" + Str(Val(Text16) + Val(Text1)) + "','" + BUYnum + "','" + Text8 + "'," _
                & "'" + gUserno + "','" + BUYdate + "')")

MsgBox "����ɹ���"
Unload Me
ExitIC

Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
'��ȡ���б��
On Error GoTo errhandle
Dim i As Integer
'�ж�IC���Ƿ�׼����
If Not InitICcard Then
    ExitIC
    Exit Sub
End If
st = chk_102(icdev)             '�����Ƿ�Ϊ�Ϸ���
If st <> 0 Then
    MsgBox ("���ǺϷ���IC�������顣")
    Exit Sub
End If
'***************�˶�����f0f0***************************
'password(0) = &HF0
'password(1) = &HF0
Dim Apass As String
Set rst = mconn.Execute("select Apass from Sysdate")
Apass = rst.Fields(0)
'Call ToBCD(Val(Left(Apass, 2)), password(0))
'Call ToBCD(Val(Right(Apass, 2)), password(1))
st = asc_hex(Apass, password(0), 2)
If st < 0 Then
    MsgBox ("��ȡ�������")
    Exit Sub
End If
rst.Close

st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("�˶�IC�������")
    Exit Sub
End If
'������Ϣ
st = srd_102_hex(icdev, 0, 2, 5, RdType(0))
If RdType(4) <> &H41 Then
MsgBox "�˿������û���������ȷ�����û�����"
Exit Sub
End If
'��ȡ�û����
Dim idTemp(3) As String
Dim idTemp2 As String
Call BCDTo(RdType(0), idTemp(0))
Call BCDTo(RdType(1), idTemp(1))
Call BCDTo(RdType(2), idTemp(2))
Call BCDTo(RdType(3), idTemp(3))
idTemp2 = Trim(idTemp(0) & idTemp(1) & idTemp(2) & idTemp(3))
'����û���Ϣ
Set rst = mconn.Execute("select * from YHdb where y_no='" + Right(idTemp2, 7) + "'")
Text10 = rst.Fields("y_id")
Text11 = rst.Fields("y_name")
Text12 = rst.Fields("y_tel")
Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "С��" & Trim(rst.Fields("y_dong")) & "��" & Trim(rst.Fields("y_dy")) & "��Ԫ" & Trim(rst.Fields("y_hao")) & "��"
Text14 = rst.Fields("y_memo")
Text15 = rst.Fields("y_no")
rst.Close
'�ж��û����Ƿ��ڵ����ˢ����ֻ��ˢ�����ܼ���1c1d����aaH
st = srd_102_hex(icdev, 1, 6, 2, RdType(0))
If st < 0 Then
    MsgBox ("������")
    Exit Sub
End If
If RdType(0) <> &HAA And RdType(1) <> &HAA Then
    MsgBox "�˿���δ�ڵ����ˢ�����������е磬���ȵ������ˢ����"
    Exit Sub
End If
'��ȡ�������ͼ��ۼƹ��������������
Set rst = mconn.Execute("select * from WTBDdb where yb_id='" + Right(idTemp2, 7) + "' and yb_buyid=(select max(yb_buyid) from WTBDdb where yb_id='" + Right(idTemp2, 7) + "')")
Text7 = rst.Fields("yb_type")
Text16 = rst.Fields("yb_tdn")
GDnum = Str(Val(rst.Fields("yb_num")) + 1)
SkinLabel14.Caption = "���Ǹ��û���" & GDnum & "�ι��硣"
rst.Close
'�����õ����ͻ�ȡ��۵�����
Set rst = mconn.Execute("select * from WTDdb where Ds_name='" + Text7 + "'")
Text2 = rst.Fields("Ds_price")
Text3 = rst.Fields("Ds_gznum")
Text4 = rst.Fields("Ds_tz")
Text6 = Format(CDate(Now), "yyyy-mm-dd")
rst.Close

Frame2.Enabled = True
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd

End Sub

Private Sub Text1_Change()
Text8 = Format(Val(Text1) * Val(Text2), "####.##")
End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9)
End Sub


Private Sub Text9_Change()
If Text9 = "" Then
Exit Sub
Else
SkinLabel17.Caption = "����" & Str(Val(Text9) - Val(Text8))
End If
End Sub
Private Sub Text9_LostFocus()
If Text9 = "" Then
Text9 = "0"
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
End Sub


