VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmJieTi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ˮ������"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8700
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "�޸�ˮ��"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmJieTi.frx":0000
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ  ��"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ǰδ��ͨ����ˮ�ۣ�"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   8295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "FrmJieTi.frx":05D3
         TabIndex        =   9
         Top             =   1440
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJieTi.frx":0711
         TabIndex        =   10
         Top             =   960
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "FrmJieTi.frx":084F
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ý���ˮ��"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   3480
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmJieTi.frx":08F5
      Top             =   3720
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   3480
      OleObjectBlob   =   "FrmJieTi.frx":0B29
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmJieTi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset

Private Sub Command1_Click()
On Error GoTo errhandle
'**********������Ϣ*******************
If MsgBox("�޸Ľ���ˮ�۽�Ӱ��֮�����й�ˮ�����������ú��������ˮ������������ã�" + Chr(13) + "��ȷ��Ҫ�޸���", vbYesNo, "���棡����") = vbNo Then
Exit Sub
End If

'**********����/ͣ�ý���ˮ��**********
If JTYes Then   '�����ǰ״̬Ϊ��������ˮ�ۣ���˴�Ϊ�ر�
    mconn.Execute ("update sysjt set jietiyesno='no'")
    JTYes = False
Else            '��������ˮ��
'����Ľ���ˮ���Ƿ�Ϸ�
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
        MsgBox "������Ŀ������д��������"
        Exit Sub
    End If
    If Val(Text3) <> Val(Text5) Then
    MsgBox "����д�������������飡"
    Text3.SetFocus
    Exit Sub
    End If
mconn.Execute ("update sysjt set jietiyesno='yes',jia1='" + Text1 + "',jia2='" + Text4 + "',jia3='" + Text6 + "',nian1='" + Text2 + "',nian2='" + Text3 + "'")
JTYes = True
End If
MsgBox "���óɹ�������"
Unload Me
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo errhandle
'**********������Ϣ*******************
If MsgBox("�޸Ľ���ˮ�۽�Ӱ��֮�����й�ˮ�����������ú��������ˮ������������ã�" + Chr(13) + "��ȷ��Ҫ�޸���", vbYesNo, "���棡����") = vbNo Then
Exit Sub
End If

'����Ľ���ˮ���Ƿ�Ϸ�
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
        MsgBox "������Ŀ������д��������"
        Exit Sub
    End If
    If Val(Text3) <> Val(Text5) Then
    MsgBox "����д�������������飡"
    Text3.SetFocus
    Exit Sub
    End If
mconn.Execute ("update sysjt set jietiyesno='yes',jia1='" + Text1 + "',jia2='" + Text4 + "',jia3='" + Text6 + "',nian1='" + Text2 + "',nian2='" + Text3 + "'")
JTYes = True
MsgBox "�޸ĳɹ�����"



Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
'**********��ǰ�Ƿ��Ѿ���ͨ����ˮ��**********
If JTYes Then
    Frame1.Caption = "��ǰ�Ѿ���ͨ����ˮ�ۣ�"
    Command1.Caption = "ͣ�ý���ˮ��"
    Command3.Visible = True
    '��䵱ǰ����ˮ������
    Set rst = mconn.Execute("select * from Sysjt")
    Text1 = rst.Fields("jia1")
    Text2 = rst.Fields("nian1")
    Text3 = rst.Fields("nian2")
    Text4 = rst.Fields("jia2")
    Text5 = rst.Fields("nian2")
    Text6 = rst.Fields("jia3")
    rst.Close
    
Else
    JTYes = False
    Frame1.Caption = "��ǰδ��ͨ����ˮ�ۣ�"
    Command1.Caption = "���ý���ˮ��"
    Command3.Visible = False
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
 If KeyAscii = 46 Then   'С����
    KeyAscii = 0
    MsgBox "ֻ��Ϊ������"
    Exit Sub
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
 If KeyAscii = 46 Then   'С����
    KeyAscii = 0
    MsgBox "ֻ��Ϊ������"
    Exit Sub
End If
End Sub

Private Sub Text3_LostFocus()
Text5 = Text3
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
 If KeyAscii = 46 Then   'С����
    KeyAscii = 0
    MsgBox "ֻ��Ϊ������"
    Exit Sub
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub

