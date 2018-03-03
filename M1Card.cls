VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "M1Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim akey(6) As Byte
Dim bkey(6) As Byte
Dim Snr As Long
Dim data32 As String * 32
Dim databuff32 As String * 32
Dim rvalue As Long
Dim wvalue As Long
Dim cardmode As Integer
Dim loadmode As Integer
Dim sector As Integer
Dim address As Integer
Dim tagtype As Integer
Dim Size As Integer

Const port As Integer = 0 '端口
Const baud As Long = 9600 '115200 '通讯波特率9600～115200
Const dataGroup As String = "0000000000000000000000000" '00" '00000"
Const CardNoLen As Integer = 7 '卡号位数

Const PROGRAMERR As String = "RF999"

Const RFINITERR As String = "RF001"
Const RFREQUESTERR As String = "RF002"
Const RFANTICOLLERR As String = "RF003"
Const RFSELECTERR As String = "RF004"
Const RFLOADKEY_A_ERR As String = "RF005"
Const RFLOADKEY_B_ERR As String = "RF006"
Const RFAUTHENTICATION_A_ERR As String = "RF007"
Const RFAUTHENTICATION_B_ERR As String = "RF008"
Const RFREADERR As String = "RF009"
Const RFWRITEERR As String = "RF010"
Const RFCHANGEB3ERR As String = "RF011"
Const RFWRITEINVERR As String = "RF012"
Const RFSETTIMEERR As String = "RF013"
Const RFSETDISPMODE As String = "RF014"
Const RFCARDIDLENGTH7 As String = "RF015"

Const RFREADCARDIDERR As String = "RF016"
Const RFREADCHARGEERR As String = "RF017"
Const RFREADIGERR As String = "RF018"

Const RFWRITECARDIDERR As String = "RF019"
Const RFWRITECHARGEERR As String = "RF020"
Const RFWRITEIGERR As String = "RF021"

Const RFCARDIDMATCH As String = "RF022"

Const RFCHECKWRITECARDIDERR As String = "RF023"
Const RFCHECKWRITECHARGEERR As String = "RF024"
Const RFCHECKWRITEIGERR As String = "RF025"

Const RFCARDNOERR       As String = "RF026"
Const RFCHARGEEQUERR       As String = "RF027"
Const RFOK As String = "OPSUCCESS"
Dim ret As String
Dim dInCharge As Double
'Dim dInIg As Long

Public Function SetDate(ByVal strDateWeekTime As String) As String
    'YYWWMMDDHHMMSS 年星期月日时分秒 每个2位09060321101010    09年星期6 3月21日 10点10分10秒
    On Error GoTo Err
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        SetDate = RFINITERR '设备初始化端口失败
        Exit Function
    End If

    st = rf_settimehex(icdev, strDateWeekTime)
    If st <> 0 Then
        SetDate = RFSETTIMEERR '设置时间失败
        Exit Function
    End If

    st = rf_disp_mode(icdev, &H1)
    If st <> 0 Then
        SetDate = RFSETDISPMODE '设置显示模式失败
        Exit Function
    End If
    SetDate = RFOK
    Call quit
    Exit Function
Err:
    Call quit
    SetDate = PROGRAMERR
    Exit Function
End Function

Private Function InitCard() As String
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        InitCard = RFINITERR '设备初始化端口失败
        Exit Function
    End If
    '寻卡
    st = rf_reset(icdev, 10) '射频读写模块复位
     If st <> 0 Then
        InitCard = RFREQUESTERR '寻卡请求失败
        Exit Function
    End If
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        InitCard = RFREQUESTERR '寻卡请求失败
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        InitCard = RFANTICOLLERR '卡防冲突，返回卡的序列号
        Exit Function
    End If

    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        InitCard = RFSELECTERR '从多个卡中选取一个给定序列号的卡
        Exit Function
    End If
    InitCard = RFOK
End Function
Private Function CheckA()
    '装载密码A
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        CheckA = RFLOADKEY_A_ERR '装载密码A出错
        Exit Function
     End If
    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        CheckA = RFAUTHENTICATION_A_ERR 'A密码验证错误
        Exit Function
    End If
    CheckA = RFOK
End Function
Private Function CheckB()
'装载B密码
    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          CheckB = RFLOADKEY_B_ERR '装载密码出错_B
          Exit Function
    End If
    '验证密码B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          CheckB = RFAUTHENTICATION_B_ERR '该卡片不属于本系统所使用
          Exit Function
    End If
    CheckB = RFOK
End Function
Private Function WriteCardNo(ByVal strCardNo As String) As String
    '写数据，卡号
    address = 4
    data32 = dataGroup & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        WriteCardNo = RFWRITECARDIDERR '写卡号失败"
        Exit Function
    End If
    st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    If st <> 0 Then
        WriteCardNo = RFWRITECARDIDERR '写卡号失败"
        Exit Function
    End If
    WriteCardNo = RFOK
End Function
Private Function WriteCharge(ByVal dCharge As Double) As String
    '写余额 2位小数
    dCharge = dCharge * 100 '2 ^ 2 '左移2位
    wvalue = CLng(dCharge)
    address = 5
    st = rf_initval(ByVal icdev, address, wvalue)
    If st <> 0 Then
        WriteCharge = RFWRITECHARGEERR '写余额失败
        Exit Function
    End If
    WriteCharge = RFOK
End Function
Private Function WriteIg(ByVal lIg As Long) As String
    '写积分
    address = 6
    st = rf_initval(ByVal icdev, address, lIg)
    If st <> 0 Then
        WriteIg = RFWRITEIGERR '写积分失败
        Exit Function
    End If
    WriteIg = RFOK
End Function
Private Function ReadCardNo(strCardNo As String) As String
    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         ReadCardNo = RFREADCARDIDERR '读取卡号错误
         Exit Function
    End If
    strCardNo = Right(databuff32, CardNoLen)
    ReadCardNo = RFOK
End Function
Private Function ReadCharge(dCharge As Double) As String
    address = 5
    st = rf_readval(ByVal icdev, address, rvalue)
    If st <> 0 Then
        ReadCharge = RFREADCHARGEERR '读取余额错误
        Exit Function
    End If
    dCharge = CDbl(rvalue)
    dCharge = dCharge / 100
    ReadCharge = RFOK
End Function
Private Function ReadIg(lIg As Long) As String
    address = 6
    st = rf_readval(ByVal icdev, address, rvalue)
    If st <> 0 Then
        ReadIg = RFREADIGERR '读取积分错误
        Exit Function
    End If
    lIg = rvalue
    ReadIg = RFOK
End Function
Public Function PutCard(ByVal strCardNo As String, ByVal dCharge As Double, ByVal dIg As Long) As String  '发卡
    On Error GoTo Err
    If Len(strCardNo) <> CardNoLen Then
        PutCard = RFCARDIDLENGTH7 '卡号位数要求
        GoTo Err
    End If

    ret = InitCard()
    If ret <> RFOK Then
        PutCard = ret
        GoTo Err
    End If
    '装载密码A
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        PutCard = RFLOADKEY_A_ERR '装载密码A出错
        GoTo Err
     End If
    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = RFAUTHENTICATION_A_ERR 'A密码验证错误
        GoTo Err
    End If

    ret = WriteCardNo(ByVal strCardNo)
    If ret <> RFOK Then
        PutCard = ret
        GoTo Err
    End If
    ret = WriteCharge(ByVal dCharge)
    If ret <> RFOK Then
        PutCard = ret
        GoTo Err
    End If
    ret = WriteIg(ByVal dIg)
    If ret <> RFOK Then
        PutCard = ret
        GoTo Err
    End If

    '改密码
    akey(0) = &H0
    akey(1) = &H2
    akey(2) = &H3
    akey(3) = &H4
    akey(4) = &H5
    akey(5) = &H6

    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5

    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        PutCard = RFCHANGEB3ERR '改密码失败
       GoTo Err
    End If

    st = rf_beep(icdev, 5)

    PutCard = RFOK
    '取消设备
    Call quit
    Exit Function
Err:
    Call quit
End Function
Public Function WriteChargeEn(ByVal dCharge As Double, ByVal dChargeBak As Double) As String
    '写余额 2位小数
    ret = WriteCharge(ByVal dCharge)
    If ret <> RFOK Then
        WriteChargeEn = RFWRITECHARGEERR
        Exit Function
    End If
    ret = ReadCharge(dInCharge)
    If ret <> RFOK Then
        ret = WriteCharge(ByVal dChargeBak)
        If ret <> RFOK Then
            WriteChargeEn = "CMT|反读反写出错，trans提交。correct data: C:" & CStr(dCharge)
            Exit Function
        Else
            WriteChargeEn = "ROL|反读反写成功，积分未写，trans回滚。"
            Exit Function
        End If
    End If
    If dCharge <> dChargeBak Or dCharge <> dInCharge Or dChargeBak <> dInCharge Then
        '反写余额
        ret = WriteCharge(ByVal dChargeBak)
        If ret <> RFOK Then
            WriteChargeEn = "CMT|校验错误，反写出错，trans提交。correct data: C:" & CStr(dCharge)
            Exit Function
        Else
            WriteChargeEn = "ROL|校验错误，反写成功，trans回滚。"
            Exit Function
        End If
    End If

    WriteChargeEn = RFOK
End Function
'函数说明：
'有卡号则判断卡号，余额大于零则写余额，积分大于零则写积分
Public Function WriteCard(ByVal strCardNo As String, ByVal dCharge As Double, ByVal dChargeBak As Double, ByVal dIg As Long) As String '写卡
    On Error GoTo Err
    ret = InitCard()
    If ret <> RFOK Then
        WriteCard = ret
        GoTo Err
    End If

    ret = CheckB()
    If ret <> RFOK Then
        WriteCard = ret
        GoTo Err
    End If

    If Len(strCardNo) > 0 Then
        Dim strInCardNo As String
        ret = ReadCardNo(strInCardNo)

        If st <> RFOK Then
            WriteCard = ret '读取卡号错误
            GoTo Err
        End If

        If strCardNo <> strInCardNo Then
            WriteCard = RFCARDNOERR '卡号不匹配
            GoTo Err
        End If
    End If

    '写积分
    If dIg > 0 Then
        ret = WriteIg(ByVal dIg)
            If ret <> RFOK Then
            WriteCard = ret
            GoTo Err
        End If
    End If

    If dCharge > 0 Then
        If dCharge <> dChargeBak Then
            WriteCard = RFCHARGEEQUERR
            GoTo Err
        Else
            ret = WriteChargeEn(ByVal dCharge, ByVal dChargeBak)
            If ret <> RFOK Then
                WriteCard = ret
                st = rf_beep(icdev, 10)
                GoTo Err
            End If
        End If
    End If
    st = rf_beep(icdev, 5)
    WriteCard = RFOK
    '取消设备
    Call quit
    Exit Function
Err:
    Call quit
End Function


Public Function ReadCard(strCardNo As String, dCharge As Double, dIg As Long) As String   '读卡
    On Error GoTo Err
     '初始化端口
    ret = InitCard()
    If ret <> RFOK Then
        ReadCard = ret
        GoTo Err
    End If

    ret = CheckB()
    If ret <> RFOK Then
        ReadCard = ret
        GoTo Err
    End If

    ret = ReadCardNo(strCardNo)
    If ret <> RFOK Then
        ReadCard = ret
        GoTo Err
    End If

    ret = ReadCharge(dCharge)
    If ret <> RFOK Then
        ReadCard = ret
        GoTo Err
    End If
    ret = ReadIg(dIg)
    If ret <> RFOK Then
        ReadCard = ret
        GoTo Err
    End If

    st = rf_beep(icdev, 5)

    ReadCard = RFOK
    '取消设备
    Call quit
    Exit Function
Err:
    Call quit
End Function

'员工卡EMP0000
Public Function EmpPutCard(ByVal strCardNo As String) As String   '发卡
    If Len(strCardNo) <> 4 Then
        EmpPutCard = RFCARDIDLENGTH7 '卡号需为4位
        Exit Function
    End If
     '初始化端口
    On Error GoTo Err
    ret = InitCard()

    '装载密码A
    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    loadmode = 0
    sector = 1

    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        EmpPutCard = RFLOADKEY_A_ERR '装载密码A出错
        GoTo Err
     End If

    '验证密码A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        EmpPutCard = RFAUTHENTICATION_A_ERR 'A密码验证错误
        GoTo Err
    End If


    '写数据，卡号
    address = 4
    data32 = "0000000000000000000000000" & "EMP" & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        EmpPutCard = RFWRITEERR '写卡号失败"
        GoTo Err
    End If


    '改密码

    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5

    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        EmpPutCard = RFCHANGEB3ERR '改密码失败
        GoTo Err
    End If

    st = rf_beep(icdev, 5)

    EmpPutCard = RFOK
    '取消设备
    Call quit
    Exit Function
Err:
    Call quit
End Function
'读取员工卡
Public Function EmpReadCard(strCardNo As String) As String    '读卡
     '初始化端口
    On Error GoTo Err
    ret = InitCard()
    ret = CheckB()
    If ret <> RFOK Then
        EmpReadCard = ret
        GoTo Err
    End If

    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         EmpReadCard = RFREADERR '读取卡号错误
         GoTo Err
    End If

    strCardNo = Right(databuff32, 4)

    st = rf_beep(icdev, 5)

    EmpReadCard = RFOK
    '取消设备
    Call quit
    Exit Function
Err:
    Call quit
End Function

Public Function RecycleCard() As String    '卡回收
    On Error GoTo Err
    ret = InitCard()

    ret = CheckB()
    If ret <> RFOK Then
        RecycleCard = ret
        GoTo Err
    End If


    akey(0) = &H0
    akey(1) = &H1
    akey(2) = &H2
    akey(3) = &H3
    akey(4) = &H4
    akey(5) = &H5

    bkey(0) = &H0
    bkey(1) = &H1
    bkey(2) = &H2
    bkey(3) = &H3
    bkey(4) = &H4
    bkey(5) = &H5
     st = rf_changeb3(icdev, sector, akey(0), 0, 0, 0, 1, 0, bkey(0))
     If st <> 0 Then
        RecycleCard = RFCHANGEB3ERR '改密码失败
        GoTo Err
    End If

    st = rf_beep(icdev, 5)

    RecycleCard = RFOK
    '取消设备
    Call quit
    Exit Function
Err:
    Call quit
End Function

