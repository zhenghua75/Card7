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

Const port As Integer = 0 '�˿�
Const baud As Long = 9600 '115200 'ͨѶ������9600��115200
Const dataGroup As String = "0000000000000000000000000" '00" '00000"
Const CardNoLen As Integer = 7 '����λ��

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
    'YYWWMMDDHHMMSS ����������ʱ���� ÿ��2λ09060321101010    09������6 3��21�� 10��10��10��
    On Error GoTo Err
    icdev = rf_init(port, baud)
    If icdev < 0 Then
        SetDate = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If
    
    st = rf_settimehex(icdev, strDateWeekTime)
    If st <> 0 Then
        SetDate = RFSETTIMEERR '����ʱ��ʧ��
        Exit Function
    End If
    
    st = rf_disp_mode(icdev, &H1)
    If st <> 0 Then
        SetDate = RFSETDISPMODE '������ʾģʽʧ��
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
        InitCard = RFINITERR '�豸��ʼ���˿�ʧ��
        Exit Function
    End If
    'Ѱ��
    st = rf_reset(icdev, 10) '��Ƶ��дģ�鸴λ
     If st <> 0 Then
        InitCard = RFREQUESTERR 'Ѱ������ʧ��
        Exit Function
    End If
    st = rf_request(icdev, 1, tagtype)
    If st <> 0 Then
        InitCard = RFREQUESTERR 'Ѱ������ʧ��
        Exit Function
    End If

    st = rf_anticoll(icdev, 0, Snr)
    If st <> 0 Then
        InitCard = RFANTICOLLERR '������ͻ�����ؿ������к�
        Exit Function
    End If
            
    st = rf_select(icdev, Snr, Size)
    If st <> 0 Then
        InitCard = RFSELECTERR '�Ӷ������ѡȡһ���������кŵĿ�
        Exit Function
    End If
    InitCard = RFOK
End Function
Private Function CheckA()
    'װ������A
    akey(0) = &HB6
    akey(1) = &H48
    akey(2) = &HA7
    akey(3) = &HF3
    akey(4) = &H2
    akey(5) = &H1C
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        CheckA = RFLOADKEY_A_ERR 'װ������A����
        Exit Function
     End If
    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        CheckA = RFAUTHENTICATION_A_ERR 'A������֤����
        Exit Function
    End If
    CheckA = RFOK
End Function
Private Function CheckB()
'װ��B����
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    loadmode = 4
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, bkey(0))
    If st <> 0 Then
          CheckB = RFLOADKEY_B_ERR 'װ���������_B
          Exit Function
    End If
    '��֤����B
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
          CheckB = RFAUTHENTICATION_B_ERR '�ÿ�Ƭ�����ڱ�ϵͳ��ʹ��
          Exit Function
    End If
    CheckB = RFOK
End Function
Private Function WriteCardNo(ByVal strCardNo As String) As String
    'д���ݣ�����
    address = 4
    data32 = dataGroup & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        WriteCardNo = RFWRITECARDIDERR 'д����ʧ��"
        Exit Function
    End If
    st = rf_check_writehex(ByVal icdev, Snr, loadmode, address, data32)
    If st <> 0 Then
        WriteCardNo = RFWRITECARDIDERR 'д����ʧ��"
        Exit Function
    End If
    WriteCardNo = RFOK
End Function
Private Function WriteCharge(ByVal dCharge As Double) As String
    'д��� 2λС��
    dCharge = dCharge * 100 '2 ^ 2 '����2λ
    wvalue = CLng(dCharge)
    address = 5
    st = rf_initval(ByVal icdev, address, wvalue)
    If st <> 0 Then
        WriteCharge = RFWRITECHARGEERR 'д���ʧ��
        Exit Function
    End If
    WriteCharge = RFOK
End Function
Private Function WriteIg(ByVal lIg As Long) As String
    'д����
    address = 6
    st = rf_initval(ByVal icdev, address, lIg)
    If st <> 0 Then
        WriteIg = RFWRITEIGERR 'д����ʧ��
        Exit Function
    End If
    WriteIg = RFOK
End Function
Private Function ReadCardNo(strCardNo As String) As String
    address = 4
    st = rf_read_hex(ByVal icdev, address, databuff32)
    If st <> 0 Then
         ReadCardNo = RFREADCARDIDERR '��ȡ���Ŵ���
         Exit Function
    End If
    strCardNo = Right(databuff32, CardNoLen)
    ReadCardNo = RFOK
End Function
Private Function ReadCharge(dCharge As Double) As String
    address = 5
    st = rf_readval(ByVal icdev, address, rvalue)
    If st <> 0 Then
        ReadCharge = RFREADCHARGEERR '��ȡ������
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
        ReadIg = RFREADIGERR '��ȡ���ִ���
        Exit Function
    End If
    lIg = rvalue
    ReadIg = RFOK
End Function
Public Function PutCard(ByVal strCardNo As String, ByVal dCharge As Double, ByVal dIg As Long) As String  '����
    On Error GoTo Err
    If Len(strCardNo) <> CardNoLen Then
        PutCard = RFCARDIDLENGTH7 '����λ��Ҫ��
        GoTo Err
    End If
    
    ret = InitCard()
    If ret <> RFOK Then
        PutCard = ret
        GoTo Err
    End If
    'װ������A
    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &HC6
    akey(3) = &H8C
    akey(4) = &HD9
    akey(5) = &HE5
    loadmode = 0
    sector = 1
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        PutCard = RFLOADKEY_A_ERR 'װ������A����
        GoTo Err
     End If
    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        PutCard = RFAUTHENTICATION_A_ERR 'A������֤����
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

    '������
    akey(0) = &HB6
    akey(1) = &H48
    akey(2) = &HA7
    akey(3) = &HF3
    akey(4) = &H2
    akey(5) = &H1C
    
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    
    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        PutCard = RFCHANGEB3ERR '������ʧ��
       GoTo Err
    End If
    
    st = rf_beep(icdev, 5)

    PutCard = RFOK
    'ȡ���豸
    Call quit
    Exit Function
Err:
    Call quit
End Function
Public Function WriteChargeEn(ByVal dCharge As Double, ByVal dChargeBak As Double) As String
    'д��� 2λС��
    ret = WriteCharge(ByVal dCharge)
    If ret <> RFOK Then
        WriteChargeEn = RFWRITECHARGEERR
        Exit Function
    End If
    ret = ReadCharge(dInCharge)
    If ret <> RFOK Then
        ret = WriteCharge(ByVal dChargeBak)
        If ret <> RFOK Then
            WriteChargeEn = "CMT|������д������trans�ύ��correct data: C:" & CStr(dCharge)
            Exit Function
        Else
            WriteChargeEn = "ROL|������д�ɹ�������δд��trans�ع���"
            Exit Function
        End If
    End If
    If dCharge <> dChargeBak Or dCharge <> dInCharge Or dChargeBak <> dInCharge Then
        '��д���
        ret = WriteCharge(ByVal dChargeBak)
        If ret <> RFOK Then
            WriteChargeEn = "CMT|У����󣬷�д������trans�ύ��correct data: C:" & CStr(dCharge)
            Exit Function
        Else
            WriteChargeEn = "ROL|У����󣬷�д�ɹ���trans�ع���"
            Exit Function
        End If
    End If
    
    WriteChargeEn = RFOK
End Function
'����˵����
'�п������жϿ��ţ�����������д�����ִ�������д����
Public Function WriteCard(ByVal strCardNo As String, ByVal dCharge As Double, ByVal dChargeBak As Double, ByVal dIg As Long) As String 'д��
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
            WriteCard = ret '��ȡ���Ŵ���
            GoTo Err
        End If
    
        If strCardNo <> strInCardNo Then
            WriteCard = RFCARDNOERR '���Ų�ƥ��
            GoTo Err
        End If
    End If
    
    'д����
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
    'ȡ���豸
    Call quit
    Exit Function
Err:
    Call quit
End Function


Public Function ReadCard(strCardNo As String, dCharge As Double, dIg As Long) As String   '����
    On Error GoTo Err
     '��ʼ���˿�
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
    'ȡ���豸
    Call quit
    Exit Function
Err:
    Call quit
End Function

'Ա����EMP0000
Public Function EmpPutCard(ByVal strCardNo As String) As String   '����
    If Len(strCardNo) <> 4 Then
        EmpPutCard = RFCARDIDLENGTH7 '������Ϊ4λ
        Exit Function
    End If
     '��ʼ���˿�
    On Error GoTo Err
    ret = InitCard()
    
    'װ������A
    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &HC6
    akey(3) = &H8C
    akey(4) = &HD9
    akey(5) = &HE5
    
    loadmode = 0
    sector = 1
       
    st = rf_load_key(ByVal icdev, loadmode, sector, akey(0))
    If st <> 0 Then
        EmpPutCard = RFLOADKEY_A_ERR 'װ������A����
        GoTo Err
     End If

    '��֤����A
    st = rf_authentication(ByVal icdev, loadmode, sector)
    If st <> 0 Then
        EmpPutCard = RFAUTHENTICATION_A_ERR 'A������֤����
        GoTo Err
    End If
    

    'д���ݣ�����
    address = 4
    data32 = "0000000000000000000000000" & "EMP" & strCardNo
    st = rf_write_hex(ByVal icdev, address, ByVal data32)
    If st <> 0 Then
        EmpPutCard = RFWRITEERR 'д����ʧ��"
        GoTo Err
    End If
    
    
    '������
    
    akey(0) = &HB6
    akey(1) = &H48
    akey(2) = &HA7
    akey(3) = &HF3
    akey(4) = &H2
    akey(5) = &H1C
    
    bkey(0) = &HC0
    bkey(1) = &H3F
    bkey(2) = &H55
    bkey(3) = &H91
    bkey(4) = &HEB
    bkey(5) = &H8
    
    st = rf_changeb3(icdev, sector, akey(0), 3, 3, 3, 3, 0, bkey(0))
     If st <> 0 Then
        EmpPutCard = RFCHANGEB3ERR '������ʧ��
        GoTo Err
    End If
    
    st = rf_beep(icdev, 5)

    EmpPutCard = RFOK
    'ȡ���豸
    Call quit
    Exit Function
Err:
    Call quit
End Function
'��ȡԱ����
Public Function EmpReadCard(strCardNo As String) As String    '����
     '��ʼ���˿�
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
         EmpReadCard = RFREADERR '��ȡ���Ŵ���
         GoTo Err
    End If
    
    strCardNo = Right(databuff32, 4)
    
    st = rf_beep(icdev, 5)
    
    EmpReadCard = RFOK
    'ȡ���豸
    Call quit
    Exit Function
Err:
    Call quit
End Function

Public Function RecycleCard() As String    '������
    On Error GoTo Err
    ret = InitCard()
    
    ret = CheckB()
    If ret <> RFOK Then
        RecycleCard = ret
        GoTo Err
    End If


    akey(0) = &HA3
    akey(1) = &HD4
    akey(2) = &HC6
    akey(3) = &H8C
    akey(4) = &HD9
    akey(5) = &HE5

    bkey(0) = &HB0
    bkey(1) = &H1B
    bkey(2) = &H4C
    bkey(3) = &H49
    bkey(4) = &HA3
    bkey(5) = &HD3
     st = rf_changeb3(icdev, sector, akey(0), 0, 0, 0, 1, 0, bkey(0))
     If st <> 0 Then
        RecycleCard = RFCHANGEB3ERR '������ʧ��
        GoTo Err
    End If
   
    st = rf_beep(icdev, 5)
    
    RecycleCard = RFOK
    'ȡ���豸
    Call quit
    Exit Function
Err:
    Call quit
End Function
