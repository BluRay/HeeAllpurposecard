VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmChaCard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ѯ��"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "FrmChaCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6285
   StartUpPosition =   2  '��Ļ����
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmChaCard.frx":030A
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1080
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   120
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmChaCard.frx":053E
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   2280
      OleObjectBlob   =   "FrmChaCard.frx":0EF2
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   615
      Left            =   600
      OleObjectBlob   =   "FrmChaCard.frx":0F53
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ  ��"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ  ��"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   600
      X2              =   5400
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "FrmChaCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RdType(0) As Byte
Dim i As Integer
Dim Para(2) As Byte
Dim oldpass As String * 4
Dim password(1) As Byte

Private Sub Command1_Click()
On Error GoTo errhandle
'�ж�IC���Ƿ�׼����
If Not InitICcard Then
    ExitIC
    Exit Sub
End If
'�����Ƿ�Ϊ�Ϸ���
st = chk_102(icdev)
If st <> 0 Then
    MsgBox ("���ǺϷ���IC�������顣")
    Exit Sub
End If
'***************�˶�����f0f0***************************
'password(0) = &HF0
'password(1) = &HF0
oldpass = "f0f0"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("�˶�IC�������")
    Exit Sub
End If
'������û����Ļ�������ʾ
'st = srd_102_hex(icdev, 2, 3, 1, RdType(0))
'If RdType(0) = &H10 Then
'MsgBox "�˿�Ϊ�û����������������á���ȷ�ϴ˿������ϣ������忨��"
'Exit Sub
'End If
'**************�忨   ******************************
'**************������ַ******************************
st = ser_102(icdev, 2, 0, 63)
If st < 0 Then
    MsgBox ("����������")
    Exit Sub
End If
Dim Test(62) As Byte
For i = 0 To 62
Test(i) = &HFF
Next i
st = swr_102_hex(icdev, 2, 0, 63, Test(0))
If st < 0 Then
  MsgBox ("�忨ʧ�ܣ���")
  Exit Sub
End If

'**************����0��******************************
st = ser_102(icdev, 0, 18, 5)
If st < 0 Then
    MsgBox ("����������")
    Exit Sub
End If
'д��ϵͳ����־
Para(0) = &H98
st = swr_102_hex(icdev, 0, 21, 1, Para(0))
If st < 0 Then
MsgBox "д��ʧ�ܣ�"
Exit Sub
End If
''*************д����־***********************
Para(0) = &H40
st = swr_102_hex(icdev, 0, 18, 1, Para(0))
If st < 0 Then
MsgBox "д��ʧ�ܣ�"
Exit Sub
End If

'2��������λ����
'*************������λ��0,�˶�����ǰ���ܶ�Ӧ����2���ж�����*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("������λ�������")
  Exit Sub
End If

'*************��������************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("���Ŀ��������")
    Exit Sub
End If
MsgBox "��ѯ�������ɹ���"

ExitIC
Unload Me
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

End Sub
