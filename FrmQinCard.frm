VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmQinCard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�忨"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "FrmQinCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6435
   StartUpPosition =   2  '��Ļ����
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5880
      OleObjectBlob   =   "FrmQinCard.frx":030A
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1800
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmQinCard.frx":053E
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   855
      Left            =   600
      OleObjectBlob   =   "FrmQinCard.frx":0DEB
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ  ��"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ  ��"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   2640
      OleObjectBlob   =   "FrmQinCard.frx":0E8F
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmQinCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Para(63) As Byte
Dim i As Integer
Dim oldpass As String * 4
Dim password(1) As Byte
Private Sub Command1_Click()
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
'***************�˶�����1b6c***************************
'password(0) = &HF0
'password(1) = &HF0 ���ٴκ˶��û�����      ????
oldpass = "1b6c"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    oldpass = "9898"        '����ʼ����
    st = asc_hex(oldpass, password(0), 2)
    st = csc_102(icdev, 2, password(0))
    If st < 0 Then
        oldpass = "f0f0"        '�¿�
        st = asc_hex(oldpass, password(0), 2)
        st = csc_102(icdev, 2, password(0))
        If st = 0 Then
        MsgBox "�˿�Ϊ�¿��������忨"
        Exit Sub
        ElseIf st < 0 Then
        MsgBox ("�˶�IC����������տ�ʧ�ܣ�")
        Exit Sub
        End If
    End If
End If
'**************����2����ַ******************************
st = ser_102(icdev, 2, 0, 64)
If st < 0 Then
    MsgBox ("�����������տ�ʧ�ܣ�")
    Exit Sub
End If
Screen.MousePointer = vbHourglass

For i = 0 To 62
Para(i) = &HFF
Next i

st = swr_102_hex(icdev, 2, 1, 63, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ����տ�ʧ�ܣ�")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'**************����1����ַ******************************
st = ser_102(icdev, 1, 0, 64)
If st < 0 Then
    MsgBox ("����ʧ�ܣ����տ�ʧ�ܣ�")
    Exit Sub
End If
Screen.MousePointer = vbHourglass

'For i = 0 To 62
'Para(i) = &HFF
'Next i

st = swr_102_hex(icdev, 1, 1, 63, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ����տ�ʧ�ܣ�")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'**************����0����ַ******************************
st = ser_102(icdev, 0, 2, 8)
If st < 0 Then
    MsgBox ("����ʧ�ܣ����տ�ʧ�ܣ�")
    Exit Sub
End If
Screen.MousePointer = vbHourglass

st = swr_102_hex(icdev, 0, 2, 8, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ����տ�ʧ�ܣ�")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'*************��������***********************
password(0) = &HF0
password(1) = &HF0
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("���Ŀ�����������տ�ʧ�ܣ�")
    Exit Sub
End If

'******************************************
    MsgBox "IC����ʼ���ɹ���"
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd

End Sub
