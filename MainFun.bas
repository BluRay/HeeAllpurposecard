Attribute VB_Name = "MainFun"
'********************************************************************************
'2008-6-17,���Ӵ������ð湦�ܣ�HY�Ǳ߰汾������Ҫ��һЩ���´ι�ȥ�ǵð����߹�������һ�¡�mdifrmsys�иĶ����ݿ��иĶ�
'2008-6-19,�����������ð湦�ܣ�HY�Ǳ߰汾������Ҫ��һЩ���´ι�ȥ�ǵð����߹�������һ�¡�mdifrmsys�иĶ�
'
'
'
'
'
'
'
'
'
'********************************************************************************



Option Explicit
Global gbDBOpenFlag As Boolean
Global YHMod As Boolean
Global GYHcha As Boolean
Global SysMod As Boolean
Global YHModS As Boolean    '���û�ά�����빺ˮ��ϸ��Ϣ
Global mconn As New ADODB.Connection                 '���ݿ�������
Global Const gPwdHS = "desktop"                      'Դ���ݿ�����
Global Const Smy = "47843182013107422770"
Global Miwen As String, Skey As String

Global JTYes As Boolean    '�Ƿ�ͨ����ˮ��

'***********************************************************************
'ȫ�� ͨѶ�ڲ�������
'***********************************************************************
Global gUserno As String            '��ǰϵͳ����Ա����
Global gUsername As String          '��ǰϵͳ����Ա
Global gPassword As String          '��ǰ����Ա����
Global gUserpower As String        '��ǰϵͳ����ԱȨ��(��������)
Global gUserOpFun As String        '��ǰϵͳ����Աְ��
Global SorD As String

Global st As Integer
Global status As Integer
Global commport As Integer              'ͨѶ��
Global icdev As Long                    '��д���ķ��س���
Global Const IC_BAUD = 57600 '9600 '
Global Const IC_BAUD2 = 9600 '

Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'************************************************************
'ICFUN
'************************************************************
Declare Function add_s Lib "mwic32.dll" (ByVal i%) As Integer
Declare Function auto_init Lib "mwic_32.dll" (ByVal port%, ByVal baud As Long) As Long
Declare Function ic_init Lib "mwic_32.dll" (ByVal port%, ByVal baud As Long) As Long
Declare Function get_status Lib "mwic_32.dll" (ByVal icdev As Long, card_S As Integer) As Integer
Declare Function turn_on Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Declare Function turn_off Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Declare Function ic_exit% Lib "mwic_32.dll" (ByVal icdev As Long)
Declare Function hex_asc% Lib "mwic_32.dll" (ByRef ascbyte As Byte, ByVal hexstr As String, lenth As Integer)
'***********************    operate at88sc102    ************************
Declare Function srd_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, ByVal offset As Integer, ByVal le As Integer, ByVal data_buffer$) As Integer
Declare Function srd_102_hex Lib "mwic_32.dll" Alias "srd_102" (ByVal icdev As Long, ByVal zone As Integer, ByVal offset As Integer, ByVal le As Integer, ByRef data_buff As Byte) As Integer

Declare Function swr_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, ByVal offset As Integer, ByVal le As Integer, ByVal data_buffer$) As Integer
Declare Function swr_102_hex Lib "mwic_32.dll" Alias "swr_102" (ByVal icdev As Long, ByVal zone As Integer, ByVal offset As Integer, ByVal le As Integer, ByRef data_buffer As Byte) As Integer

Declare Function ser_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, ByVal offset As Integer, ByVal le As Integer) As Integer

Declare Function csc_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Declare Function rsc_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Declare Function wsc_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Declare Function rsct_102 Lib "mwic_32.dll" (ByVal icdev As Long, counter As Integer) As Integer

Declare Function cesc_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Declare Function resc_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Declare Function wesc_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Declare Function resct_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer, counter%) As Integer

Declare Function clrpr_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer) As Integer
Declare Function clrrd_102 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal zone As Integer) As Integer

Declare Function fakefus_102 Lib "mwic_32.dll" (ByVal icdev As Long, mode%) As Integer
Declare Function blow_102 Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Declare Function chk_102 Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
'*****************
'Declare Function get_status Lib "MWIC_32.dll" (ByVal icdev As Long, card_S As Integer) As Integer

Declare Function set_baud Lib "mwic_32.dll" (ByVal icdev As Long, ByVal baud As Long) As Integer

Declare Function cmp_dvsc Lib "mwic_32.dll" (ByVal icdev As Long, ByVal length As Integer, ByVal data_buffer$) As Integer
Declare Function srd_dvsc Lib "mwic_32.dll" (ByVal icdev As Long, ByVal length As Long, ByVal data_buffer$) As Integer
Declare Function swr_dvsc Lib "mwic_32.dll" (ByVal icdev As Long, ByVal length As Integer, ByVal data_buffer$) As Integer
Declare Function setsc_md Lib "mwic_32.dll" (ByVal icdev As Long, ByVal mode As Integer) As Integer
  


Declare Function srd_ver Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByVal data_buffer$) As Integer
Declare Function auto_pull Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Declare Function dv_beep Lib "mwic_32.dll" (ByVal icdev As Long, ByVal time As Integer) As Integer


Declare Function asc_hex Lib "mwic_32.dll" (ByVal asc$, ByRef hex As Byte, ByVal le&) As Integer

Declare Function asc_asc% Lib "mwic_32.dll" (ByVal sorc$, ByRef des As Byte, ByVal le&)

Declare Function ic_encrypt Lib "mwic_32.dll" (ByVal key As String, ByVal ptrsource As String, ByVal le As Integer, ByRef ptrdest As Byte) As Integer
Declare Function ic_decrypt Lib "mwic_32.dll" (ByVal key As String, ByRef ptrdest As Byte, ByVal le As Integer, ByVal ptrsource As String) As Integer



'************************************************************
'������OpenDatabaseX()
'���ܣ��Թ���ʽ�����ݿ�
'����ֵ�� SUCCESS--�ɹ�
'         FAILED--ʧ��
'************************************************************
Public Function OpenDatabaseX() As Boolean
On Error GoTo errhandle
mconn.Open "DSN=YHWTMS;uid=sa;pwd=" + gPwdHS + ""
OpenDatabaseX = True

Exit Function
errhandle:
OpenDatabaseX = False
MsgBox (Error(ErR))
Call QuitSystem
End Function
'************************************************************
'* �˳�ϵͳ
'************************************************************
Public Sub QuitSystem()
 On Error GoTo errhandle
    '�û��˳�,�ر����ݿ�
    If gbDBOpenFlag = True Then
   ' ���²�����¼
'   dateout = Format(Now, "yyyy-MM-dd hh:mm:ss")
'   mconn.Execute ("update oprecord set date_out='" + dateout + "' where date_in='" + datein + "'")
'   mconn.Close
    End If
    On Error Resume Next
    gbDBOpenFlag = False
    mconn.Close
    ExitIC

    '�ر�MDI�����Ӵ���
    'frmMDI ���屾��Ϊforms(0),�Ӵ����forms(1)--forms(forms.count-1)
    Dim i As Integer
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i
    End
Exit Sub
errhandle:
   MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
   Resume Next
End Sub

'****************************************************************
'����:���ַ���ת��Ϊָ�����ȵĸ�ʽ
'����: �� �ַ������� >= ָ������, ����right(�ַ���,ָ������)
'      �����㳤��,���ַ���ǰ����"0"
'����:��׼����.��.�յĸ�ʽ "80"--"0080","8"--"08"
'****************************************************************
Public Function FormatString(numberstr As String, formatlength As Integer) As String
On Error GoTo errhandle
    Dim i, strlen, mulstep As Integer
    
    strlen = Len(Trim$(numberstr))
    If formatlength >= strlen Then
        mulstep = formatlength - strlen
        FormatString = Trim$(numberstr)
    Else
        FormatString = Left$(Trim$(numberstr), formatlength)
        Exit Function
    End If
    For i = 1 To mulstep
        FormatString = "0" + FormatString
    Next i
Exit Function
errhandle:
   MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
   Resume Next
    
End Function

'***********************************************************
'��ʼ��IC�� ,�ɹ�����TRUE,ʧ�ܷ���FALSE
'***********************************************************
Public Function InitICcard() As Boolean
Dim status As Integer
If icdev < 0 Then       'ͨѶ��û�г�ʼ��
redo:   icdev = ic_init(commport - 1, IC_BAUD2)      '��ʼ��ͨѶ��  ���أ� >=0 ��ȷ  <0 ����
        If icdev < 0 Then                       '��ʼ������
            icdev = ic_init(commport - 1, IC_BAUD)
            
            If icdev < 0 Then
            Dim ans As Integer
            ans = MsgBox("IC����д����ʼ������,�Ƿ����³�ʼ����", 16 + vbRetryCancel, "��Ϣ����")
                If ans = vbRetry Then
                    GoTo redo
                Else
                    InitICcard = False
                    Exit Function
                End If
            End If
            
            
        End If
redo1:
        Dim st As Integer
        st = get_status(icdev, status)       '����IC����д����ǰ״̬ ���� 0 ��ȷ���ӣ�<0  �������ӻ�û�г�ʼ��;status=1 �п�����,0���޿�����
        If st = 0 And status = 0 Then
            ans% = MsgBox("û�п�����,��忨��", 48 + vbRetryCancel, "��Ϣ����")
            If ans = vbRetry Then
                GoTo redo1
            Else
                InitICcard = False
                ExitIC
                Exit Function
            End If
        End If
Else
    InitICcard = True
End If
InitICcard = True

End Function

'ȡ��IC����д�����ڵĴ���
Public Function GetCommPort() As Integer
  Screen.MousePointer = vbHourglass
  Dim i, st As Integer
  For i = 0 To 3
    icdev = ic_init(i, IC_BAUD2)
    If icdev < 0 Then
    icdev = ic_init(i, IC_BAUD)
        If icdev > 0 Then
            st = get_status(icdev, status)
            If st = 0 Then
                GetCommPort = i + 1
                Exit For
            End If
        End If
    Else
        st = get_status(icdev, status)
        If st = 0 Then
            GetCommPort = i + 1
            Exit For
        End If
        
    End If
  Next i
  If icdev < 0 Or st < 0 Then
        GetCommPort = 10
        MsgBox "ͨѶ�ڳ�ʼ��ʧ�ܣ���ȷ�϶�д�������Ӳ��򿪵�Դ��", vbOKOnly + vbCritical, "������ʾ..."
  End If
  Screen.MousePointer = vbDefault
End Function

'�ж�IC����д���Ƿ�׼����
Public Function IsICReady() As Boolean
If commport = 10 Then
    Beep
    MsgBox "IC����д��δ׼���ã�", vbOKOnly + vbInformation, App.Title
    IsICReady = False
Else
    IsICReady = True
End If
End Function


'************************************************************
'�˳�IC����д���Ĳ���,���ύͨѶ��
'************************************************************
Public Sub ExitIC()
    Dim st As Integer
    st = turn_off(icdev)     '��IC���µ�
    st = ic_exit(icdev)
    If st = 0 Then
        icdev = -1
    End If
    
End Sub
'****************************************************************
'* ����     limitnumber(KeyAscii As Integer)
'* ����:    ���Ƽ�����ַ�Ϊ�����ַ���
'*          ��Ϊ�����ַ�,���ظ��ַ���ASCII��ֵ��
'* ��ڲ���: �����ַ���ASCII��ֵ
'* ����:    ��Ϊ�����ַ�,���ظ��ַ���ASCII��ֵ��
'*          ����,����0
'****************************************************************
Public Function limitnumber(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 '����
        Case 46 'С����
        Case 8  '�˸�
        Case 27 'ESCAPE��
        Case 13 '�س�
        Case 32
            KeyAscii = 0
        Case Else
             KeyAscii = 0    'ȡ���ַ�
             Beep            '���������ź�
             MsgBox "����� ������������( 0 - 9)�� ", vbOKOnly + vbCritical, App.Title
    End Select
    limitnumber = KeyAscii
End Function

'**********��һ����λʮ������ת��ΪBCD��******************
Public Function ToBCD(num1 As String, num2 As Byte)
    num2 = Val(Left(num1, 1)) * 16 + Val(Right(num1, 1))
End Function
'**********��һ��BCD��ת��Ϊԭʮ������********************
Public Function BCDTo(num1 As Byte, num2 As String)
 num2 = FormatString(Str(num1 \ 16), 1) & FormatString(Str(num1 Mod 16), 1)
End Function


'**********ע��ϵͳ����***********************************
Public Function power(mw As String, my As String) As String
Dim i As Integer
Dim fromstr(8) As Byte
Dim midstr(8) As Byte
Dim tostr(8) As Byte
Dim jiemi As String
Dim tmp(7) As Single
For i = 0 To 7
    fromstr(i) = Val(Mid(Trim(mw), i + 1, 1))
Next i
For i = 0 To 7
    midstr(i) = Val(Mid(my, i * 2 + 1, 2))
Next i
For i = 0 To 6
    fromstr(i) = fromstr(i) Xor fromstr(i + 1)
    tmp(i) = fromstr(i)
Next i
    fromstr(7) = fromstr(7) Xor 23
    tmp(7) = fromstr(7)
For i = 0 To 7
    tostr(i) = (tmp(i) * midstr(i)) Mod 256
Next i

jiemi = ""
For i = 0 To 7
    If Len(Trim(hex(tostr(i)))) = 1 Then
        jiemi = jiemi + "0" + Trim(hex(tostr(i)))
    Else
        jiemi = jiemi + Trim(hex(tostr(i)))
    End If
Next i
power = jiemi
End Function

