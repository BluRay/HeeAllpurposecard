Attribute VB_Name = "MainFun"
'********************************************************************************
'2008-6-17,增加次数试用版功能，HY那边版本功能上要新一些，下次过去记得把两边功能整合一下。mdifrmsys有改动数据库有改动
'2008-6-19,增加日期试用版功能，HY那边版本功能上要新一些，下次过去记得把两边功能整合一下。mdifrmsys有改动
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
Global YHModS As Boolean    '从用户维护进入购水详细信息
Global mconn As New ADODB.Connection                 '数据库主连接
Global Const gPwdHS = "desktop"                      '源数据库密码
Global Const Smy = "47843182013107422770"
Global Miwen As String, Skey As String

Global JTYes As Boolean    '是否开通阶梯水价

'***********************************************************************
'全局 通讯口操作变量
'***********************************************************************
Global gUserno As String            '当前系统操作员代号
Global gUsername As String          '当前系统操作员
Global gPassword As String          '当前操作员口令
Global gUserpower As String        '当前系统操作员权限(操作级别)
Global gUserOpFun As String        '当前系统操作员职能
Global SorD As String

Global st As Integer
Global status As Integer
Global commport As Integer              '通讯口
Global icdev As Long                    '读写器的返回常数
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
'函数：OpenDatabaseX()
'功能：以共享方式打开数据库
'返回值： SUCCESS--成功
'         FAILED--失败
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
'* 退出系统
'************************************************************
Public Sub QuitSystem()
 On Error GoTo errhandle
    '用户退出,关闭数据库
    If gbDBOpenFlag = True Then
   ' 更新操作记录
'   dateout = Format(Now, "yyyy-MM-dd hh:mm:ss")
'   mconn.Execute ("update oprecord set date_out='" + dateout + "' where date_in='" + datein + "'")
'   mconn.Close
    End If
    On Error Resume Next
    gbDBOpenFlag = False
    mconn.Close
    ExitIC

    '关闭MDI所有子窗体
    'frmMDI 窗体本身为forms(0),子窗体从forms(1)--forms(forms.count-1)
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
'功能:将字符串转换为指定长度的格式
'返回: 若 字符串长度 >= 指定长度, 返回right(字符串,指定长度)
'      否则不足长度,在字符串前补足"0"
'例如:标准化年.月.日的格式 "80"--"0080","8"--"08"
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
'初始化IC卡 ,成功返回TRUE,失败返回FALSE
'***********************************************************
Public Function InitICcard() As Boolean
Dim status As Integer
If icdev < 0 Then       '通讯口没有初始化
redo:   icdev = ic_init(commport - 1, IC_BAUD2)      '初始化通讯口  返回： >=0 正确  <0 错误
        If icdev < 0 Then                       '初始化错误
            icdev = ic_init(commport - 1, IC_BAUD)
            
            If icdev < 0 Then
            Dim ans As Integer
            ans = MsgBox("IC卡读写器初始化错误,是否重新初始化？", 16 + vbRetryCancel, "信息窗口")
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
        st = get_status(icdev, status)       '返回IC卡读写器当前状态 返回 0 正确连接，<0  错误连接或没有初始化;status=1 有卡插入,0则无卡插入
        If st = 0 And status = 0 Then
            ans% = MsgBox("没有卡插入,请插卡！", 48 + vbRetryCancel, "信息窗口")
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

'取得IC卡读写器所在的串口
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
        MsgBox "通讯口初始化失败！请确认读写器已联接并打开电源。", vbOKOnly + vbCritical, "警告提示..."
  End If
  Screen.MousePointer = vbDefault
End Function

'判断IC卡读写器是否准备好
Public Function IsICReady() As Boolean
If commport = 10 Then
    Beep
    MsgBox "IC卡读写器未准备好！", vbOKOnly + vbInformation, App.Title
    IsICReady = False
Else
    IsICReady = True
End If
End Function


'************************************************************
'退出IC卡读写器的操作,并提交通讯口
'************************************************************
Public Sub ExitIC()
    Dim st As Integer
    st = turn_off(icdev)     '对IC卡下电
    st = ic_exit(icdev)
    If st = 0 Then
        icdev = -1
    End If
    
End Sub
'****************************************************************
'* 函数     limitnumber(KeyAscii As Integer)
'* 功能:    限制键入的字符为数字字符。
'*          若为数字字符,返回该字符的ASCII码值。
'* 入口参数: 键入字符的ASCII码值
'* 返回:    若为数字字符,返回该字符的ASCII码值；
'*          否则,返回0
'****************************************************************
Public Function limitnumber(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9 '数字
        Case 46 '小数点
        Case 8  '退格
        Case 27 'ESCAPE键
        Case 13 '回车
        Case 32
            KeyAscii = 0
        Case Else
             KeyAscii = 0    '取消字符
             Beep            '发出错误信号
             MsgBox "输入错！ 必须输入数字( 0 - 9)。 ", vbOKOnly + vbCritical, App.Title
    End Select
    limitnumber = KeyAscii
End Function

'**********将一个两位十进制数转化为BCD码******************
Public Function ToBCD(num1 As String, num2 As Byte)
    num2 = Val(Left(num1, 1)) * 16 + Val(Right(num1, 1))
End Function
'**********将一个BCD码转化为原十进制数********************
Public Function BCDTo(num1 As Byte, num2 As String)
 num2 = FormatString(Str(num1 \ 16), 1) & FormatString(Str(num1 Mod 16), 1)
End Function


'**********注册系统加密***********************************
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

