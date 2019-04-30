VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrderInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsOrders is the base class which has common features
' required in other classes
Private aOOTId As Long
Private aAcct As Long
Private aRSPRef As String
Private aExistingAsid As Long
Private aExistingUCLLAsid As Long
Private aStartDate As String
Private aStartTime As String
Private aOrderType As String
Private aServiceType As String
Private aTNo As String
Private aBTCode As String
Private aProvPlatform As String
Private aExchange As String
Private aSAMId As Long
Private aSpotArr As Variant
Private aWIPWOP As String
Private aIntrlve As String
Private aCustName As String
Private aCustContactNo As String
Private aCustContactNo2 As String
Private aLspRef As String
Private aLspName As String
Private aGainName As String
Private aEubaPlan As String
Private aSONo As Long
Private aCreatedOrderDetails As String
Private aErrCode As String
Private aBBorPOTS As String
Private aPOTSExch As String
Private aExchMismatch As Boolean

Public Property Get ExchMismatch() As Boolean
    ExchMismatch = aExchMismatch
End Property

Public Property Get AccountNo() As Long
    AccountNo = aAcct
End Property

Public Property Let AccountNo(ByVal AccountNo As Long)
    aAcct = AccountNo
End Property

Public Property Get ootid() As Long
    ootid = aOOTId
End Property

Public Property Let ootid(ByVal ootid As Long)
    aOOTId = ootid
End Property

Public Property Get rspRef() As String
    rspRef = aRSPRef
End Property

Public Property Get POTSExch() As String
    POTSExch = aPOTSExch
End Property

Public Property Let rspRef(ByVal rspRef As String)
    aRSPRef = rspRef
End Property

Public Property Get ExistingAsid() As Long
    ExistingAsid = aExistingAsid
End Property

Public Property Let ExistingAsid(ByVal ExistingAsid As Long)
    aExistingAsid = ExistingAsid
End Property

Public Property Get ExistingUCLLAsid() As Long
    ExistingUCLLAsid = aExistingUCLLAsid
End Property

Public Property Let ExistingUCLLAsid(ByVal ExistingUCLLAsid As Long)
    aExistingUCLLAsid = ExistingUCLLAsid
End Property

Public Property Get StartDate() As String
    StartDate = aStartDate
End Property

Public Property Let StartDate(ByVal StartDate As String)
    aStartDate = StartDate
End Property

Public Property Get StartTime() As String
    StartTime = aStartTime
End Property

Public Property Let StartTime(ByVal StartTime As String)
    aStartTime = StartTime
End Property

Public Property Get OrderType() As String
    OrderType = aOrderType
End Property

Public Property Let LspRef(ByVal LspRef As String)
    aLspRef = LspRef
End Property

Public Property Let LspName(ByVal LspName As String)
    aLspName = LspName
End Property

Public Property Let GainName(ByVal GainName As String)
    aGainName = GainName
End Property


Public Property Let OrderType(ByVal OrderType As String)
    aOrderType = OrderType
End Property
'aWIPWOP
Public Property Get WIPWOP() As String
    WIPWOP = aWIPWOP
End Property

Public Property Let WIPWOP(ByVal WIPWOP As String)
    aWIPWOP = WIPWOP
End Property

Public Property Get ServiceType() As String
    ServiceType = aServiceType
End Property

Public Property Let ServiceType(ByVal ServiceType As String)
    aServiceType = ServiceType
End Property

Public Property Get tNo() As String
    tNo = aTNo
End Property

Public Property Let tNo(ByVal tNo As String)
    aTNo = tNo
End Property

Public Property Let BTCode(ByVal BTCode As String)
    aBTCode = BTCode
End Property

Public Property Let BBorPOTS(ByVal BBorPOTS As String)
    aBBorPOTS = BBorPOTS
End Property

Public Property Get ProvPlatform() As String
    ProvPlatform = aProvPlatform
End Property

Public Property Let ProvPlatform(ByVal ProvPlatform As String)
    aProvPlatform = ProvPlatform
End Property

Public Property Get exchange() As String
    exchange = aExchange
End Property

Public Property Let exchange(ByVal exchange As String)
    aExchange = exchange
End Property

Public Property Get samID() As Long
    samID = aSAMId
End Property

Public Property Let samID(ByVal samID As Long)
    aSAMId = samID
End Property

Public Property Get SpotArr() As Variant
    SpotArr = aSpotArr
End Property

Public Property Let SpotArr(ByVal SpotArr As Variant)
    aSpotArr = SpotArr
End Property

Public Property Get Intrlve() As String
    Intrlve = aIntrlve
End Property

Public Property Let Intrlve(ByVal Intrlve As String)
    aIntrlve = Intrlve
End Property

Public Property Get CustName() As String
    CustName = aCustName
End Property

Public Property Let CustName(ByVal CustName As String)
    aCustName = CustName
End Property

Public Property Get CustContactNo() As String
    CustContactNo = aCustContactNo
End Property

Public Property Let CustContactNo(ByVal CustContactNo As String)
    aCustContactNo = CustContactNo
End Property

Public Property Get CustContactNo2() As String
    CustContactNo2 = aCustContactNo2
End Property

Public Property Let CustContactNo2(ByVal CustContactNo2 As String)
    aCustContactNo2 = CustContactNo2
End Property

Public Property Get EubaPlan() As String
    EubaPlan = aEubaPlan
End Property

Public Property Let EubaPlan(ByVal EubaPlan As String)
    aEubaPlan = EubaPlan
End Property

Public Property Get ErrCode() As String
    ErrCode = aErrCode
End Property

Public Property Let ErrCode(ByVal ErrCode As String)
    aErrCode = ErrCode
End Property

Public Property Get SoNo() As String
    SoNo = aSONo
End Property

Public Property Let SoNo(ByVal SoNo As String)
    aSONo = SoNo
End Property
'

Public Property Get CreatedOrderDetails() As String
    CreatedOrderDetails = aCreatedOrderDetails
End Property

Public Property Let CreatedOrderDetails(ByVal CreatedOrderDetails As String)
    aCreatedOrderDetails = CreatedOrderDetails
End Property

Public Sub StagetoWWC()
    Dim icms As New clsICMS, i As Integer
    With icms
        .ConnectPS
        If .onChorusMainMenu = True Then
            .setCursor 24, 58
            .SendKeyString "01"
            .EnterKey
            For i = 0 To 3
                .FunctionKey_05
                If .scrRead(9, 27, 47) = "Line Level Alert Code" Then
                    .FunctionKey_03
                End If
            Next
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub StagetoScheduleScrn()
    Dim icms As New clsICMS, i As Integer
    With icms
        .ConnectPS
        If .onChorusMainMenu = True Then
            .setCursor 24, 58
            .SendKeyString "22"
            .EnterKey
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub EnterASIDinWWC()
    Dim icms As New clsICMS, i As Integer, temStr As String
    If aExistingAsid = 0 Then
        aErrCode = "Empty Account ID"
        Set icms = Nothing
        Exit Sub
    End If
    With icms
        .ConnectPS
        temStr = .scrRead(2, 32, 49)
        If temStr <> "WORK WITH CUSTOMER" Then
            aErrCode = "Not in WWC Screen"
            Set icms = Nothing
            Exit Sub
        End If
        For i = 0 To 2
            .FunctionKey_05
        Next
        .setCursor 7, 7
        .SendKeyString aExistingAsid
        For i = 0 To 2
            .EnterKey
        Next
        .setCursor 21, 19
    End With
    temStr = vbNullString
    Set icms = Nothing
End Sub

Public Sub EnterAcctIDinWWC()
    Dim icms As New clsICMS, i As Integer, temStr As String
    If aAcct = 0 Then
        aErrCode = "Empty Account ID"
        Set icms = Nothing
        Exit Sub
    End If
    With icms
        .ConnectPS
        temStr = .scrRead(2, 32, 49)
        If temStr <> "WORK WITH CUSTOMER" Then
            aErrCode = "Not in WWC Screen"
            Set icms = Nothing
            Exit Sub
        End If
        For i = 0 To 2
            .FunctionKey_05
        Next
        .setCursor 7, 44
        .SendKeyString aAcct
        For i = 0 To 1
            .EnterKey
        Next
    End With
    temStr = vbNullString
    Set icms = Nothing
End Sub

Public Function POTSZoneSpotCode() As String
    Dim icms As New clsICMS, i As Integer, temStr As String
    If aExistingAsid = 0 Then Exit Function
    With icms
        If .ConnectPS = True Then
            If .onChorusMainMenu = True Then
                Call StagetoWWC
                .ConnectPS
                .FunctionKey_09
                .setCursor 6, 18
                .FunctionKey_04
                .FunctionKey_09
                .setCursor 20, 23
                .SendKeyString "7"
                .EnterKey
                .setCursor 4, 21
                For i = 0 To 10
                    .FieldExitKey
                Next
                .setCursor 10, 73
                .SendKeyString aSAMId
                .EnterKey
                .setCursor 13, 2
                .SendKeyString "x"
                .EnterKey
                If .scrRead(9, 27, 53) = "Regional Offers Information" Then .FunctionKey_12
                temStr = Trim(.scrRead(13, 12, 31, False))
                If (InStr(temStr, "Auckland") > 0) Or (InStr(temStr, "Wellington") > 0) Or (InStr(temStr, "Christchurch") > 0) Then
                    POTSZoneSpotCode = "3RH03"
                Else
                    POTSZoneSpotCode = "3RH01"
                End If
            End If
        End If
        For i = 0 To 3
            If .scrRead(6, 28, 50) = "CHORUS CUSTOMER SERVICE" Then Exit For
            .FunctionKey_12
        Next
    End With
    Set icms = Nothing
End Function

Public Sub StagetoOrderEnter()
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(2, 32, 49)
        If temStr <> "WORK WITH CUSTOMER" Then
            aErrCode = "Not in WWC Screen"
            Set icms = Nothing
            Exit Sub
        End If
        .setCursor 21, 19
        .SendKeyString "04"
        .EnterKey
        If .scrRead(1, 27, 53) <> "SERVICE ORDER ENTRY/ENQUIRY" Then
            .FunctionKey_12
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub EnterIntoExistingSO(ByVal SoNo As Long)
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(1, 27, 53)
        If temStr = "SERVICE ORDER ENTRY/ENQUIRY" Then
            For i = 0 To 8
                If i = 0 Or i = 8 Then .setCursor 4, 25
                If i > 0 And i < 8 Then .FieldExitKey
                If i = 8 Then
                    .SendKeyString CStr(SoNo)
                    .setCursor 10, 68: .SendKeyString "A"
                    .EnterKeywithoutWaitT
                End If
            Next
            For i = 0 To 30
                If InStr(.scrRead(24, 2, 37), "in use by") > 0 Then
                    modMiscFunc.status_update "", "", "", "This macro is not stuck, WY order is locked by either a person or one of Automation. Macro keeps retrying to access in for 30 secs before dumping it to exception list"
                    .callSleep 1200
                    .setCursor 4, 31
                    .EnterKeywithoutWaitT
                Else
                    Exit For
                End If
            Next i
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub EnterIntoExistingSOwithRetryFunction(ByVal SoNo As Long)
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(1, 27, 53)
        If temStr = "SERVICE ORDER ENTRY/ENQUIRY" Then
            For i = 0 To 8
                If i = 0 Or i = 8 Then .setCursor 4, 25
                If i > 0 And i < 8 Then .FieldExitKey
                If i = 8 Then
                    .SendKeyString CStr(SoNo)
                    .setCursor 10, 68: .SendKeyString "A"
                    .EnterKey
                End If
            Next
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub StagetoInitialSOScrn()
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(1, 27, 53)
        If temStr <> "SERVICE ORDER ENTRY/ENQUIRY" Then
            aErrCode = "Not in Service Order Entry Screen"
            Set icms = Nothing
            Exit Sub
        End If
        .FunctionKey_06
    End With
    Set icms = Nothing
End Sub

Public Sub Fill_ISOE_Screen()
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String, OpenOrder As String
    Dim RFSMonth As String, RFSDate As String, RFSTime As String, RFSYear As String
    Dim Screen_Type As String
    RFSYear = Right(aStartDate, 2)
    RFSMonth = Left(Mid(aStartDate, 3), 2)
    RFSDate = Left(aStartDate, 2)
    RFSTime = aStartTime
    With icms
        .callSleep 250
        .ConnectPS
        temStr = .scrRead(1, 28, 54)
        If temStr = "Initial Service Order Entry" Then
            .setCursor 3, 29
            .SendKeyString aOrderType
            For i = 0 To 1
                .setCursor 4, 29
                If i = 0 Then
                    .FieldExitKey
                Else
                    If Left(aTNo, 1) = "T" Then
                        temStr = Right(aTNo, Len(aTNo) - 1)
                    Else
                        temStr = aTNo
                    End If
                    .SendKeyString temStr
                    .EnterKey
                End If
            Next
            If .scrRead(16, 23, 36) = "F1 to continue" Then .FunctionKey_01
            If aOrderType <> "UR" Then
                If .scrRead(8, 2, 15) = "Account Number" Then
                    For i = 0 To 2
                        .EnterKey
                    Next i
                Else
                    If aServiceType = "AP" Then
                        .setCursor 5, 29
                        .SendKeyString "FICT"
                        .EnterKey
                        .setCursor 6, 69
                        .SendKeyString "BUBAS"
                        .EnterKey
                    ElseIf aOrderType = "T1" Then
                        .setCursor 5, 29
                        .SendKeyString "NEAX"
                        .EnterKey
                        .setCursor 6, 28
                        .SendKeyString aExistingAsid
                        .EnterKey
                        For i = 0 To 1
                            .EnterKey
                        Next
                    Else
                        .setCursor 8, 28
                        .FieldExitKey
                        .setCursor 8, 28
                        .SendKeyString aAcct: .EnterKey
                        temStr = .scrRead(24, 12, 48)
                        If InStr(temStr, "Line card attached") > 0 Then
                            aErrCode = temStr
                            .FunctionKey_03
                            .FunctionKey_03
                            Set icms = Nothing
                            Exit Sub
                        End If
                        If InStr(temStr, "exists") > 0 Then
                            aErrCode = "Open order exists for this line number"
                            .FunctionKey_03
                            .FunctionKey_03
                            Set icms = Nothing
                            Exit Sub
                        End If
                    End If
                End If
            Else
                .callSleep 500
                If aOrderType = "UR" Then
                    .setCursor 8, 37
                End If
                For i2 = 0 To 1
                    .setCursor 6, 28
                    .EnterKey
                    .callSleep 500
                Next
                If aOrderType = "UR" Then .callSleep 2000
'                temStr = .scrRead(24, 3, 74)
'                If InStr(temStr, "attached to") > 0 Then
'                    temStr = Mid(temStr, InStr(temStr, "attached to order") + 18, 50)
'                    temStr = Left(temStr, InStr(temStr, ". Press") - 1)
'                    aSONo = temStr
'                    Stop
'
'                End If
            End If
            If aOrderType = "UR" Then
                If InStr(.scrRead(24, 3, 64), "Open Service Order") > 0 Or InStr(.scrRead(24, 3, 64), "attached to") > 0 Then
                    aErrCode = "ERR_UR_CREATED_ALREADY"
                    Set icms = Nothing
                    Exit Sub
                End If
            End If
            If aServiceType = "RI" Then
                Do While Trim(.scrRead(15, 52, 78)) = ""
                    .EnterKey
                Loop
                For i = 9 To 13
                    If .scrRead(i, 1, 23) = "Schedule Code and Month" Then Exit For
                Next
                .setCursor i, 35
                .SendKeyString RFSMonth
                .setCursor i, 73
                .SendKeyString aServiceType
                .setCursor i + 1, 73
                .SendKeyString RFSDate & RFSMonth & RFSYear
                .setCursor i + 2, 73
                If RFSTime = "AM" Then
                    .SendKeyString "1159"
                ElseIf RFSTime = "PM" Then
                    .SendKeyString "1659"
                End If
                .setCursor i + 4, 73
                .SendKeyString "Y"
                .EnterKey
            ElseIf aOrderType = "UR" Then
                .callSleep 500
                .setCursor 11, 35
                .SendKeyString RFSMonth
                .setCursor 12, 73
                .SendKeyString RFSDate & RFSMonth & RFSYear
                .setCursor 13, 73
                If RFSTime = "AM" Then
                    .SendKeyString "1159"
                ElseIf RFSTime = "PM" Then
                    .SendKeyString "1659"
                End If
                .setCursor 15, 73
                .SendKeyString "Y"
            Else
                .setCursor 11, 35
                .SendKeyString RFSMonth
                .setCursor 11, 73
                .SendKeyString aServiceType
                .setCursor 12, 73
                .SendKeyString RFSDate & RFSMonth & RFSYear
                .setCursor 13, 73
                If RFSTime = "AM" Then
                    .SendKeyString "1159"
                ElseIf RFSTime = "PM" Then
                    .SendKeyString "1659"
                End If
                .setCursor 15, 73
                .SendKeyString "Y"
                .EnterKey
            End If
            If InStr(.scrRead(24, 3, 75), "override") > 0 Then
                .FunctionKey_13
                Stop
            End If
            If InStr(.scrRead(24, 3, 75), "less than the default Due Date/Time") > 0 Then
                .FunctionKey_14
                .setCursor 15, 35
                .SendKeyString tNo
                .EnterKey
                If .scrRead(1, 20, 61) = "Service Order Physical Address Information" Then
                    .FunctionKey_06
                End If
            Else
                .FunctionKey_06
            End If
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub Fill_SAM_Screen(ByVal Func_C_X_R As String)
    Dim icms As New clsICMS, i As Integer, temStr As String, ProcessedAsid As Long
    Dim month As String, aDate As String, Time As String
    If Func_C_X_R = "" Then
        aErrCode = "Emply Function Code"
        Exit Sub
    End If
    aExchMismatch = False
    With icms
        .ConnectPS
        If .scrRead(1, 20, 61) = "Service Order Physical Address Information" Then
            .setCursor 7, 28
            If .scrRead(11, 28, 35) = "LNP FICT" Then
                For i = 0 To 1
                    If i = 0 Then .FunctionKey_12
                    If i = 1 Then .FunctionKey_09
                Next
                .setCursor 20, 23
                .SendKeyString "7"
                .EnterKey
                .setCursor 10, 73
                .FieldExitKey
                .setCursor 10, 73
                .SendKeyString aSAMId
                .EnterKey
                .setCursor 13, 2
                .SendKeyString "x"
                .EnterKey
            Else
                For i = 0 To 15
                    .FieldExitKey
                Next
                If aServiceType = "RI" Then
                    .FunctionKey_12
                    .FunctionKey_09
                    .setCursor 20, 23
                    .SendKeyString "7"
                    .EnterKey
                    .setCursor 10, 73
                    .FieldExitKey
                    .setCursor 10, 73
                    .SendKeyString aSAMId
                    .EnterKey
                    .setCursor 13, 2
                    .SendKeyString "x"
                    .EnterKey
                    'VALIDATION - Need check if ucll asid match between icms and wireline
                Else
                    .setCursor 14, 72
                    .SendKeyString aSAMId: .EnterKey
                End If
            End If
            If aOrderType = "E5" Then
                .FunctionKey_10
                .FunctionKey_06
                Set icms = Nothing
                Exit Sub
            End If
            If aServiceType = "RI" Then
                ProcessedAsid = aExistingUCLLAsid
            ElseIf InStr("E5_WY_E4_WX", aOrderType) > 0 Then
                ProcessedAsid = aExistingAsid
            End If
            For i = 17 To 22
                If Trim(.scrRead(i, 2, 2)) = "X" Then
                    .setCursor i, 2
                    .FieldExitKey
                End If
                If i = 22 Then
                    .setCursor 14, 72
                    .FunctionKey_11
                    'Line number 0093798437 is not supported in the Address provided
                    .FunctionKey_09
                    .setCursor 21, 23
                    .SendKeyString "2"
                    .EnterKey
                    .setCursor 8, 19
                    .FieldExitKey
                    .setCursor 8, 19
                    .SendKeyString ProcessedAsid
                    .setCursor 8, 19
                    .EnterKey
                    If .scrRead(12, 13, 19) = Right(CStr(ProcessedAsid), 7) Then
                        If InStr("C_D_I", .scrRead(12, 27, 27)) > 0 Then
                            .setCursor 12, 2
                            .SendKeyString "X"
                            .EnterKey
                            .setCursor 17, 2
                            .SendKeyString Func_C_X_R
                            If Func_C_X_R = "R" Then
                                .FunctionKey_02
                                Exit For
                            Else
                                .EnterKey
                                Exit For
                            End If
                        End If
                    End If
                    'further code is required
                End If
                If InStr(.scrRead(24, 2, 75), "NEW PHONE") > 0 Then
                    .setCursor 24, 75
                    .SendKeyString "N"
                    .EnterKey
                    aExchMismatch = True
                End If
                If InStr(.scrRead(24, 2, 60), "TRANSFERRING") > 0 Then
                    .setCursor 24, 75
                    .SendKeyString "Y"
                    .EnterKey
                End If
                temStr = Trim(.scrRead(i, 18, 27))
                If temStr = "" Then temStr = 0
                'VALIDATION - NEED CREATE A VAR CALLED UCLLAISD
                If temStr = ProcessedAsid Then
                    .setCursor i, 2
                    .SendKeyString Func_C_X_R
                    If Func_C_X_R = "R" Or Func_C_X_R = "X" Then
                        .EnterKey
                    Else
                        .FunctionKey_02
                    End If
                    'EXCEPTION: 24,2,40 "Line Card Reserved to S/O 07941942"
                    .callSleep 500
                    If InStr(.scrRead(24, 2, 75), "NEW PHONE") > 0 Then
                        .setCursor 24, 75
                        .SendKeyString "N"
                        .EnterKey
                        aExchMismatch = True
                    End If
                    If InStr(.scrRead(24, 2, 60), "TRANSFERRING") > 0 Then
                        .setCursor 24, 75
                        .SendKeyString "Y"
                        .EnterKey
                    End If
                    Exit For
                End If
            Next i
            If InStr(.scrRead(24, 2, 74), "has no Plant at all") > 0 Then
                aErrCode = "ERR_FAILED_TO_ATTACH_LINECARD"
                .setCursor 17, 2
                .FunctionKey_10
            End If
            If InStr(.scrRead(24, 2, 75), "NEW PHONE") > 0 Then
                .setCursor 24, 75
                .SendKeyString "N"
                .EnterKey
                aExchMismatch = True
            End If
            If InStr(.scrRead(24, 2, 75), "TRANSFERRING") > 0 Then
                .setCursor 24, 75
                .SendKeyString "Y"
                .EnterKey
            End If
            For i = 0 To 5
                temStr = .scrRead(1, 23, 59)
                If temStr = "INTEGRATED CUSTOMER MANAGEMENT SYSTEM" Then Exit For
                temStr = .scrRead(1, 17, 64)
                If temStr = "Service Order Contact and Scheduling Information" Then Exit For
                .FunctionKey_06
                If .scrRead(24, 2, 37) = "DO YOU WANT TO ALLOCATE A NEW NUMBER" Then
                    .setCursor 24, 75
                    .SendKeyString "N"
                    .EnterKey
                    aExchMismatch = True
                End If
                If InStr(.scrRead(24, 2, 75), "TRANSFERRING") > 0 Then
                .setCursor 24, 75
                .SendKeyString "Y"
                .EnterKey
            End If
            Next
            For i = 0 To 1
                temStr = .scrRead(1, 17, 64)
                If temStr = "Service Order Contact and Scheduling Information" Then Exit For
                .FunctionKey_06
            Next
        End If
    End With
    Set icms = Nothing
End Sub


Public Sub Fill_Schedule_Screen()
    Dim icms As New clsICMS, i As Integer, temStr As String
    Dim month As String, aDate As String, Time As String, scheduleCode As String
    With icms
        .ConnectPS
        If .scrRead(1, 17, 64) = "Service Order Contact and Scheduling Information" Then
            'Scheduling Screen:
            .EnterKey
            .callSleep 1500
            If InStr(.scrRead(24, 2, 56), "override") > 0 Then
                If Left(aOrderType, 1) = "T" Then
                    If DASHBOARD.chkNotApplyBauCode = False Then
                        scheduleCode = Trim(.scrRead(13, 23, 29))
                        'Events for find a next available RFS Date.
                        .FunctionKey_01
                        temStr = ScheduleHandler(scheduleCode)
                        If temStr = "01011901AM" Then
                            If aErrCode = "" Then
                                aErrCode = "ERR_Unable to find a valid RFS date"
                            Else
                                aErrCode = aErrCode & vbNewLine & "ERR_Unable to find a valid RFS date"
                            End If 'aErrCode = ""
                            .FunctionKey_13: .FunctionKey_06
                        End If 'temStr = "01011901AM"
                        .ConnectPS
                        .FunctionKey_09
                        .callSleep 500
                        .setCursor 10, 74
                        .SendKeyString "EU"
                        .EnterKey
                        .setCursor 9, 28
                        .SendKeyString aStartDate
                        .setCursor 9, 57
                        If aStartTime = "AM" Then
                            .SendKeyString "1159"
                        ElseIf aStartTime = "PM" Then
                            .SendKeyString "1659"
                        End If 'aStartTime = "AM"
                        If .scrRead(12, 28, 29) <> "EU" Then
                            .setCursor 12, 28
                            .SendKeyString "EU"
                        End If
                        If .scrRead(9, 30, 31) <> .scrRead(13, 36, 37) Then
                            month = .scrRead(9, 30, 31)
                            .setCursor 13, 36
                            .SendKeyString month
                        End If
                        .setCursor 8, 28
                        .FieldExitKey
                        .setCursor 8, 28
                        If IsNumeric(aTNo) = True Then
                            .SendKeyString aTNo
                        Else
                            .SendKeyString Right(aTNo, Len(aTNo) - 1)
                        End If
                        .EnterKey
                        .FunctionKey_02
                    Else
                        If InStr(.scrRead(24, 2, 56), "override") > 0 Then .FunctionKey_13
                    End If 'DASHBOARD.chkNotApplyBauCode = False
                Else
                    If InStr(.scrRead(24, 2, 56), "override") > 0 Then .FunctionKey_13
                End If
            Else
                If InStr(.scrRead(24, 2, 56), "override") > 0 Then .FunctionKey_13
            End If
            For i = 0 To 5
                temStr = .scrRead(1, 20, 60)
                If temStr = "Service Order Additional Line Information" Then Exit For
                .FunctionKey_06
                If InStr(.scrRead(24, 2, 56), "override") > 0 Then .FunctionKey_13
            Next
            For i = 0 To 5
                temStr = .scrRead(1, 20, 60)
                If temStr = "Service Order Additional Line Information" Then Exit For
                .FunctionKey_06
            Next
            If aServiceType = "AP" Then
                .setCursor 20, 14
                .SendKeyString aOOTId
                .setCursor 21, 14
                .SendKeyString "OO&T"
                .EnterKey
            End If
            For i = 0 To 3
                If .scrRead(1, 32, 35) = "SPOT" Then Exit For
                .FunctionKey_06
                If InStr(.scrRead(24, 2, 72), "Listing Code must be defined") > 0 Then
                    .setCursor 7, 21
                    .SendKeyString "UL"
                    .FunctionKey_06
                End If
            Next
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub Fill_SPOTS_Screen()
    Dim icms As New clsICMS, i As Integer, i2 As Integer, i_spots As Integer, temStr As String
    Dim spotsArr() As String
    Dim month As String, aDate As String, Time As String
    With icms
        .ConnectPS
        i_spots = 0
        i2 = 9
        If aServiceType <> "AP" Then
            For i2 = 9 To 20
                If Trim(.scrRead(i2, 4, 8, True)) <> "" Then
                    .setCursor i2, 2
                    .SendKeyString "d"
                    .EnterKey
                Else
                    Exit For
                End If
            Next i2
        End If
        If aOrderType = "UR" Then
            .setCursor i2, 4
            .SendKeyString aSpotArr
            .EnterKey
        Else
            For i = i2 To i2 + UBound(aSpotArr)
                If aSpotArr(i_spots) = "" Then Exit For
                .setCursor i, 4
                .SendKeyString aSpotArr(i_spots)
                If aSpotArr(i_spots) = "3IREC" Then
                    .setCursor i, 67
                    .SendKeyString "SAL"
                    .EnterKey
                End If
                .EnterKey
                i_spots = i_spots + 1
                If aServiceType = "RI" Or aOrderType = "E5" Then
                    aWIPWOP = "WIP"
                ElseIf aOrderType = "WY" Then
                    aWIPWOP = "WOP"
                End If
                If InStr(.scrRead(24, 1, 67), "Enter Sale Term or press F13") > 0 Then
                    .setCursor i, 67
                    .SendKeyString aWIPWOP
                    .EnterKey
                End If
            Next
        End If
        'Out Spots
        If aOrderType = "UR" Then
            For i = 16 To 21
                temStr = .scrRead(i, 16, 38)
                If InStr(temStr, "MPF monthly ren") > 0 Then
                    .setCursor i, 2
                    .SendKeyString "O"
                    .EnterKey
                    Exit For
                End If
            Next
            .SendKeyString aStartDate
            .EnterKey
        End If
        .setCursor 9, 2
        .FunctionKey_23
        For i = 10 To 15
            If i = 15 Then .FunctionKey_12
            If .scrRead(i, 2, 2) <> "" And .scrRead(i, 10, 10) = "I" Then
                .setCursor i, 15
                .SendKeyString aOOTId
                .setCursor i, 29
                .SendKeyString aRSPRef
                .EnterKey
            End If
        Next
        .FunctionKey_06
        .callSleep 500
        'Interleaving
        If .scrRead(2, 27, 54) = "Component Value Input Screen" Then
            If aServiceType = "RI" Then
                .setCursor 9, 47 '---Correction please urgently
                .SendKeyString aGainName 'Gaining Service Provider ID
                '.setCursor 10, 47
                '.SendKeyString "11247" 'Gaining Service Provider ID
                .setCursor 11, 47 '---Correction please urgently
                .SendKeyString aLspName 'Losing Service Provider ID 'CORRECTION - NEED TYPE IN-POT LSP REF
                .setCursor 12, 47
                .SendKeyString aCustName 'CONTNAME Customer Contact Name
                .setCursor 13, 47
                .SendKeyString aCustContactNo 'CONTNAME Customer Contact Number
                .setCursor 14, 47
                If aLspRef = "0" Then
                    .SendKeyString ""
                Else
                    .SendKeyString aLspRef
                End If
                'CORRECTION - NEED TYPE IN-POT LSP REF(Corrected on 17th)
            ElseIf aServiceType = "AP" Then
                .setCursor 8, 47
                .FieldExitKey
                .setCursor 8, 47
                .SendKeyString aIntrlve
                .EnterKey
            End If
            .FunctionKey_06
            .callSleep 1000
            For i = 0 To 5
                If InStr(.scrRead(22, 2, 40), "Press ENTER to continue") > 0 Then
                    .setCursor 22, 2
                    .EnterKeywithoutWaitT
                    .callSleep 500
                End If
            Next
            If .scrRead(1, 23, 57) = "Service Order Directory Information" Then
                .FunctionKey_06
            End If
            If .scrRead(2, 31, 50) = "Work With Call Plans" Then
                .FunctionKey_06
            End If
            For i = 0 To 5
                If InStr(.scrRead(22, 2, 40), "Press ENTER to continue") > 0 Then
                    .setCursor 22, 2
                    .EnterKeywithoutWaitT
                    .callSleep 500
                End If
            Next
            For i = 0 To 5
                If .scrRead(1, 31, 39) = "Memo Type" Then Exit For
                .FunctionKey_06
            Next
        End If
    End With
    Set icms = Nothing
End Sub



Public Sub Fill_Memo_Screen()
    Dim icms As New clsICMS, i As Integer, i_spots As Integer, temStr As String
    Dim spotsArr() As String
    Dim month As String, aDate As String, Time As String
    With icms
        .ConnectPS
        If InStr(.scrRead(22, 2, 40), "Press ENTER to continue") > 0 Then
            .setCursor 9, 2
            .EnterKey
        End If
        If InStr(.scrRead(1, 23, 57), "Service Order Directory Information") > 0 Then
            .FunctionKey_06
        End If
        If InStr(.scrRead(2, 31, 50), "Work With Call Plans") > 0 Then
            .FunctionKey_06
        End If
        If .scrRead(1, 31, 34) = "Memo" Then
            If aOrderType = "UR" Then
                .setCursor 3, 2: .SendKeyString "UCLL to ADSL Migration"
                .setCursor 4, 2: .SendKeyString "Create Order Macro: Created a clearing Order for OOT Request " & aOOTId
                .setCursor 5, 2: .SendKeyString "NOTE: Linecard will be posted off at " & aStartTime & _
                                                " on " & aStartDate
                .EnterKey
            ElseIf aServiceType = "RI" Then
                .setCursor 3, 2
                .SendKeyString SETTING.Range("B12").value
                .setCursor 3, 2
                For i = 4 To 22
                    If .scrRead(i, 2, 15) = "OOT ID REQUEST" Then
                        .setCursor i, 17
                        .SendKeyString CStr(aOOTId)
                    End If
                    If .scrRead(i, 16, 41) = "Cust Contact Name & Number" Then
                        .setCursor i, 2
                        .SendKeyString CStr(aCustName) & " " & CStr(aCustContactNo) & ":Cust Contact Name & Number"
                    End If
                    If aCustContactNo2 <> "" Then
                        If .scrRead(i, 16, 33) = "Cust Mobile Number" Then
                            .setCursor i, 2
                            .SendKeyString aCustContactNo2 & ":Cust Mobile Number"
                        End If
                    End If
                Next
            Else
                .setCursor 3, 2
                .FunctionKey_08
                .setCursor 3, 2
                .SendKeyString "I2"
                .EnterKey
                .setCursor 3, 2
                .SendKeyString SETTING.Range("B14").value
                .setCursor 4, 2
                .SendKeyString "Request ID: " & aOOTId
                .EnterKey
                For i = 4 To 22
                    If InStr(.scrRead(i, 2, 79), ":Cust Contact Name") > 0 Then
                        .setCursor i, 2
                        .SendKeyString aCustName & " - " & aCustContactNo
                    End If
'                    If InStr(.scrRead(i, 2, 79), ":Cust Phone Number") > 0 Then
'                        .setCursor i, 2
'                        .SendKeyString aCustContactNo
'                    End If
                    If aCustContactNo2 <> "0" Then
                        If InStr(.scrRead(i, 2, 79), ":Cust Phone Number") > 0 Then
                            .setCursor i, 2
                            .SendKeyString aCustContactNo2
                        End If
                    End If
                    If InStr(.scrRead(i, 2, 79), "eg FS/40K, FS/90K") > 0 Then
                        .setCursor i, 2
                        .SendKeyString aEubaPlan & " 0 " & aIntrlve
                    End If
                Next
                .EnterKey
            End If
            .FunctionKey_06
            If .scrRead(4, 3, 25) = "Press Enter to Continue" Then .EnterKey
            If aOrderType = "UR" Then
                .setCursor 18, 64
                .SendKeyString "17"
                .EnterKey
                .setCursor 14, 29
                .SendKeyString Right(aTNo, Len(aTNo) - 1)
                .setCursor 15, 29
                .SendKeyString "000000"
                .setCursor 19, 29
                .FunctionKey_08
                .setCursor 19, 72
                .SendKeyString aStartDate
                .EnterKey
                If .scrRead(24, 2, 46) = "Service Order Line Card has not been attached" Then
                    aErrCode = "ERR_UR_ORDER_POST_IN_ERROR"
                    .FunctionKey_01
                    Set icms = Nothing
                    Exit Sub
                End If
                .FunctionKey_23
                .FunctionKey_01
            End If
            If aOrderType = "E5" Then
                .FunctionKey_09
                .setCursor 21, 52
                .SendKeyString "6S"
                .EnterKey
                .FunctionKey_12
            End If
            If InStr(.scrRead(5, 2, 51), "OPEN was issued") > 0 Then
                .setCursor 20, 18
                For i = 0 To 5
                    If InStr(.scrRead(5, 2, 51), "OPEN was issued") < 1 Then
                        Exit For
                    Else
                        .FunctionKey_12
                        .callSleep 250
                    End If
                Next
                If .scrRead(1, 29, 52) = "Service Order Completion" Then
                    .setCursor 19, 29
                    .FunctionKey_01
                    .callSleep 250
                End If
            End If
            aSONo = Trim(.scrRead(5, 13, 19, True))
            temStr = .scrRead(4, 13, 23): temStr = Replace(temStr, "-", "")
            aCreatedOrderDetails = aSONo & "," & temStr
        End If
    End With
    Set icms = Nothing
End Sub

Public Sub BacktoChorusMenu()
    Dim icms As New clsICMS
    With icms
        If .ConnectPS = True Then
            Do While .scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE"
                If .scrRead(2, 32, 49) = "WORK WITH CUSTOMER" Then
                    .setCursor 21, 19
                    .FunctionKey_05
                End If
                .FunctionKey_03
            Loop
        End If
    End With
End Sub

Public Sub ChangetoBTCode()
    Dim icms As New clsICMS, i As Integer, temStr As String
    If DASHBOARD.chkNotApplyBauCode = True Then
    If aBTCode = "" Then
        aErrCode = "ERR_BTCODE_FAILED_UPDATE"
        Set icms = Nothing
        Exit Sub
    End If
    With icms
        .ConnectPS
        temStr = .scrRead(1, 23, 56)
        If temStr <> "Service Order Entry / Enquiry Menu" Then
            aErrCode = "Failed to update BT Code"
            Set icms = Nothing
            Exit Sub
        End If
        .setCursor 18, 64
        .SendKeyString "14"
        .EnterKey
        For i = 0 To 1
            .setCursor 13, 23
            If i = 0 Then
                .FieldExitKey
            Else
                .SendKeyString aBTCode
                .EnterKey
                .callSleep 500
                If InStr(.scrRead(24, 2, 56), "override") > 0 Then
                    .FunctionKey_13
                End If
                temStr = Trim(.scrRead(10, 27, 34))
                temStr = Format(Replace(temStr, "/", ""), "000000")
                If .scrRead(10, 36, 40) = "11:59" Then
                    temStr = temStr & "AM"
                ElseIf .scrRead(10, 36, 40) = "16:59" Then
                    temStr = temStr & "PM"
                End If
                If aStartDate & aStartTime <> temStr Then
                    .setCursor 13, 23
                    .FunctionKey_09
                    .setCursor 10, 74
                    .SendKeyString "EU"
                    .EnterKey
                    .setCursor 8, 28
                    If Left(aTNo, 1) = "T" Then aTNo = Replace(aTNo, "T", "")
                    .SendKeyString aTNo
                    .setCursor 9, 28
                    .SendKeyString aStartDate
                    .setCursor 9, 57
                    If aStartTime = "AM" Then
                        .SendKeyString "1159"
                    ElseIf aStartTime = "PM" Then
                        .SendKeyString "1659"
                    End If
                    .setCursor 10, 28
                    .SendKeyString "Y"
                    If .scrRead(9, 30, 31) <> .scrRead(13, 36, 37) Then
                        temStr = .scrRead(9, 30, 31)
                        .setCursor 13, 36
                        .SendKeyString temStr
                    End If
                    .EnterKey
                    If InStr(.scrRead(24, 2, 56), "override") > 0 Then
                        .FunctionKey_13
                    End If
                End If
            End If
        Next
        .callSleep 500
        .FunctionKey_01
    End With
    Set icms = Nothing
    End If
End Sub

Public Sub StagetoSalesOrderScreen()
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(1, 27, 53)
        If temStr <> "SERVICE ORDER ENTRY/ENQUIRY" Then
            aErrCode = "Not in SO Screen"
            Set icms = Nothing
            Exit Sub
        End If
        .setCursor 4, 25
        .FunctionKey_05
    End With
    Set icms = Nothing
End Sub

Public Sub StageanOrdertoXL()
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(1, 23, 56)
        If temStr <> "Service Order Entry / Enquiry Menu" Then
            aErrCode = "Not in SO Screen"
            Set icms = Nothing
            Exit Sub
        End If
        .setCursor 18, 64
        .FunctionKey_09
        .setCursor 21, 52
        .SendKeyString "XL"
        .EnterKey
        .FunctionKey_12
    End With
    Set icms = Nothing
End Sub

Public Function SOM_New_Order(ByVal PriOrder As String, othrOrder As String, Optional SecOrder As String) As Long
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String, contAttachOrder As Boolean
    If aAcct = 0 Then
        aErrCode = "Invalid Acct No"
        Exit Function
    End If
    contAttachOrder = True
    With icms
        .ConnectPS
        temStr = .scrRead(1, 30, 52)
        If temStr <> "SALES ORDER MAINTENANCE" Then
            aErrCode = "Not in SOM Screen"
            Set icms = Nothing
            Exit Function
        End If
        .FunctionKey_06
        For i = 0 To 6
            .FieldExitKey
        Next
        .setCursor 4, 18
        .SendKeyString aAcct
        .setCursor 6, 32
        .SendKeyString aStartDate
        .setCursor 6, 47
        If aStartTime = "AM" Then
            .SendKeyString "1159"
        ElseIf aStartTime = "PM" Then
            .SendKeyString "1659"
        End If
        .EnterKey
        .callSleep 150
        .setCursor 6, 70
        If InStr(.scrRead(24, 2, 74), "Due date is invalid") > 0 Then
            If aErrCode = "" Then
                aErrCode = "ERR_INVALID_DUE_DATE"
                SOM_New_Order = 0
                .FunctionKey_12
                Do While .scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE"
                    .FunctionKey_03
                Loop
                Set icms = Nothing
                Exit Function
            End If
        End If
        SOM_New_Order = Trim(.scrRead(3, 68, 78, True))
        .callSleep 350
        .FunctionKey_07
        If PriOrder <> "" Then
            .setCursor 4, 32
            For i = 0 To 1
                .FieldExitKey
            Next
            .setCursor 4, 32
            .SendKeyString PriOrder
            .setCursor 5, 32
            .SendKeyString "A"
            .EnterKey
            .setCursor 8, 2
            .SendKeyString "x"
            .EnterKeywithoutWaitT
            For i = 0 To 15
                If InStr(.scrRead(25, 2, 73), "another user") > 0 Then
                    modMiscFunc.status_update "", "", "", "This macro is not stuck, BB order is locked by either a person or one of Automation. It keeps retrying to access in for 30 secs before dumping it to exception list"
                    .callSleep 2000
                    For i2 = 0 To 1
                        .setCursor 8, 2
                        If i2 = 0 Then
                            .FieldExitKey
                        Else
                            .SendKeyString "x"
                        End If
                    Next i2
                    .EnterKeywithoutWaitT
                Else
                    contAttachOrder = True
                    Exit For
                End If
            Next
            modMiscFunc.status_update "", "", "", "Attached WY order onto SOM"
            If .scrRead(7, 35, 44, False) = "INACTIVE" Then
                .FunctionKey_12
                contAttachOrder = False
            End If
            If contAttachOrder = True Then
                temStr = .scrRead(7, 8, 9)
                .setCursor 18, 64
                .SendKeyString "10"
                .EnterKey
                .setCursor 3, 2
                .FunctionKey_08
                .setCursor 3, 2
                .SendKeyString "I1"
                .EnterKey
                .setCursor 3, 2
                .SendKeyString "ICMS SOM#" & SOM_New_Order
                .EnterKey
                If temStr = "6S" Then
                    .setCursor 3, 2
                    .FunctionKey_08
                    .setCursor 3, 2
                    .SendKeyString "I1"
                    .EnterKey
                    .setCursor 3, 2
                    .SendKeyString SETTING.Range("B15").value
                    .EnterKey
                End If
                .FunctionKey_01
                .setCursor 18, 64
                .SendKeyString "17"
                .EnterKey
                .FunctionKey_01
                .FunctionKey_12
                .FunctionKey_01
                .EnterKey
            Else 'contAttachOrder = True
                aErrCode = "ERR_Failed to attach orders into SOM"
                .FunctionKey_01
            End If 'contAttachOrder = True
        End If 'PriOrder <> ""
        contAttachOrder = False
        If SecOrder <> "0" Then
            temStr = .scrRead(9, 10, 17)
            .setCursor 7, 71
            .FunctionKey_07
            .setCursor 4, 32
            For i = 0 To 1
                .FieldExitKey
            Next
            .setCursor 4, 32
            .SendKeyString SecOrder
            .setCursor 5, 32
            .SendKeyString "A"
            .EnterKey
            .FunctionKey_01
            If .scrRead(22, 2, 16) = "Link to primary" Then
                .setCursor 22, 22
                '.SendKeyString temStr
                .EnterKey
            End If
            If InStr(.scrRead(24, 2, 77), "updated by another user") > 0 Then
                If aErrCode = "" Then
                    aErrCode = "ERR_Failed to attach order to ICMS SOM"
                Else
                   aErrCode = aErrCode & vbNewLine & "ERR_Failed to attach order to ICMS SOM"
                End If
            End If
            For i = 9 To 13
                temStr = .scrRead(i, 67, 67)
                If temStr = "" Then Exit For
                If temStr = "S" Then
                    .setCursor i, 5
                    .SendKeyString "P"
                    .EnterKey
                    If InStr(.scrRead(22, 2, 44), "ENTER") > 0 Then .EnterKey
                End If
            Next
            Call BacktoChorusMenu
            Call StagetoWWC
            'Call EnterASIDinWWC
            Call StagetoOrderEnter
            Call EnterIntoExistingSO(SecOrder)
            .ConnectPS
            .SendKeyString "10"
            .EnterKey
            .setCursor 3, 2
            .FunctionKey_08
            .setCursor 3, 2
            .SendKeyString "I1"
            .EnterKey
            .setCursor 3, 2
            .SendKeyString "ICMS SOM#" & SOM_New_Order
            .EnterKey
            .FunctionKey_01
            .FunctionKey_12
        End If
        If othrOrder = "" Then temStr = "noOtherOrder"
        If othrOrder = "0" Then temStr = "noOtherOrder"
        If temStr <> "noOtherOrder" Then
            .setCursor 7, 71
            .FunctionKey_07
            .setCursor 4, 32
            For i = 0 To 1
                .FieldExitKey
            Next
            .setCursor 4, 32
            .SendKeyString othrOrder
            .setCursor 5, 32
            .SendKeyString "A"
            .EnterKey
            .FunctionKey_01
            If .scrRead(22, 2, 16) = "Link to primary" Then
                'temStr = .scrRead(9, 10, 17)
                .setCursor 22, 22
                '.SendKeyString temStr
                .EnterKey
            End If
            If InStr(.scrRead(24, 2, 77), "updated by another user") > 0 Then
                If aErrCode = "" Then
                    aErrCode = "ERR_Failed to attach order to ICMS SOM"
                Else
                   aErrCode = aErrCode & vbNewLine & "ERR_Failed to attach order to ICMS SOM"
                End If
            End If
        End If
        Call BacktoChorusMenu
    End With
    Set icms = Nothing
End Function

Public Sub SOM_Attach_Order(ByVal Work_Order As Long)
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        temStr = .scrRead(6, 70, 73)
        If temStr <> "OPEN" Then
            aErrCode = "Not a valid SOM"
            Set icms = Nothing
            Exit Sub
        End If
        temStr = .scrRead(1, 30, 52)
        If temStr <> "SALES ORDER MAINTENANCE" Then
            aErrCode = "Not in SOM Screen"
            Set icms = Nothing
            Exit Sub
        End If
        .FunctionKey_06
        For i = 0 To 6
            If i = 0 Or i = 6 Then .setCursor 4, 18
            .FieldExitKey
        Next
        .SendKeyString aAcct
        .setCursor 6, 32
        .SendKeyString aStartDate
        .setCursor 6, 47
        .SendKeyString aStartTime
        .EnterKey
        .setCursor 7, 71
        .FunctionKey_07
        For i = 0 To 1
            .FieldExitKey
            .setCursor 4, 32
        Next
        .SendKeyString Work_Order
        .setCursor 5, 32
        .SendKeyString "O"

    End With
    Set icms = Nothing
End Sub
Function CheckUCLLExch(ByVal ucllASID As Long) As String
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        'temStr = .scrRead(
        If .onChorusMainMenu = True Then
            For i = 0 To 1
                .setCursor 24, 58
                If i = 0 Then
                    .SendKeyString "10"
                Else
                    .SendKeyString "02"
                End If
                .EnterKey
            Next i
            For i = 0 To 1
                .setCursor 17, 60
                If i = 0 Then
                    .FunctionKey_10
                    .setCursor 4, 28
                    .SendKeyString CStr(ucllASID)
                    .EnterKey
                ElseIf i = 1 Then
                    .SendKeyString "3"
                    .EnterKey
                    temStr = .scrRead(8, 53, 57)
                    If InStr(temStr, "-") > 0 Then
                        temStr = Left(temStr, InStr(temStr, "-") - 1)
                    ElseIf InStr(temStr, "/") > 0 Then
                        temStr = Left(temStr, InStr(temStr, "/") - 1)
                    End If
                End If
            Next i
            .FunctionKey_01
            .FunctionKey_03
            .setCursor 24, 58
            .SendKeyString "98"
            .EnterKey
            CheckUCLLExch = temStr
        End If

    End With
    temStr = vbNullString
    Set icms = Nothing
End Function


Function CheckPOTSType(ByVal POTS As Long) As String
    Dim icms As New clsICMS, i As Integer, temStr As String
    With icms
        .ConnectPS
        'temStr = .scrRead(
        If .onChorusMainMenu = True Then
            .setCursor 24, 58
            .SendKeyString "01"
            .EnterKey
            For i = 0 To 2
                .FunctionKey_05
            Next
            .setCursor 7, 7
            .SendKeyString Str(POTS)
            .EnterKey
            If .scrRead(24, 2, 43) = "Line number not found in Subscriber Master" Then
                CheckPOTSType = "RTB"
                .FunctionKey_03
                Set icms = Nothing
                Exit Function
            Else
                For i = 0 To 5
                    If .scrRead(6, 28, 50) = "CHORUS CUSTOMER SERVICE" Then Exit For
                    .FunctionKey_12
                Next
            End If
            For i = 0 To 1
                Select Case i
                    Case 0
                        temStr = "08"
                    Case 1
                        temStr = "01"
                End Select
                .setCursor 24, 58
                .SendKeyString temStr
                .EnterKey
            Next i
            If .scrRead(2, 29, 50) = "National Number Scroll" Then
                .setCursor 6, 26
                .SendKeyString Str(POTS)
                .EnterKey
            End If
            temStr = Trim(.scrRead(5, 26, 30, True))
            If temStr = "" Or InStr(temStr, "RTB") > 0 Then
                CheckPOTSType = "RTB"
            Else
                CheckPOTSType = "SPARK"
            End If
            aPOTSExch = modMiscFunc.RegexReplace(temStr, "\d", "")
        End If
        .FunctionKey_03
        .setCursor 24, 58
        .SendKeyString "98"
        .EnterKey
    End With
    temStr = vbNullString
    Set icms = Nothing
End Function

Function CheckZoneBeforeEnterOrder() As String
    Dim icms As New clsICMS, i As Integer, tempStr As String
    If aSAMId = 0 Then
        aErrCode = "Invalid SAM ID"
        Exit Function
    End If
    With icms
        If .onWWCMenu = True Then
            .setCursor 21, 19
            .FunctionKey_09
            If .scrRead(2, 24, 55) = "Regional Offers Customer Enquiry" Then
                .setCursor 6, 18
                .FunctionKey_04
                If .scrRead(2, 29, 54) = "Service Area Master Scroll" Then
                    .FunctionKey_09
                    .setCursor 20, 23
                    .SendKeyString "07"
                    .EnterKey
                    For i = 0 To 25
                        .FieldExitKey
                    Next
                    .setCursor 10, 73
                    .SendKeyString aSAMId
                    .EnterKey
                    .setCursor 13, 2
                    .SendKeyString "x"
                    .EnterKey
                    tempStr = .scrRead(13, 12, 28)
                    If (InStr(tempStr, "Auckland") > 0) Or _
                        (InStr(tempStr, "Christchurch") > 0) Or _
                        (InStr(tempStr, "Wellinton") > 0) Then
                        CheckZoneBeforeEnterOrder = "AWC"
                    Else
                        CheckZoneBeforeEnterOrder = "RON"
                    End If
                End If
                .FunctionKey_03
            End If
        End If
        
    End With
    Set icms = Nothing
End Function

Sub DealAutoCreatedWY(ByVal SoNo As Long)
    Dim icms As New clsICMS, i As Integer, temStr As String
    If aExistingAsid = 0 Then Exit Sub
    With icms
        If .ConnectPS = True Then
            If .onChorusMainMenu = True Then
                Call StagetoWWC
                'Call EnterASIDinWWC
                Call StagetoOrderEnter
                Call EnterIntoExistingSO(SoNo)
                .callSleep 500
                .ConnectPS
                For i = 0 To 30
                    If i = 30 Then
                        If aErrCode = "" Then
                            aErrCode = "ERR_FAILED_TO_UPDATE_WY_ORDER"
                        Else
                            aErrCode = aErrCode & vbNewLine & "ERR_FAILED_TO_UPDATE_WY_ORDER"
                        End If
                    End If
                    If InStr(.scrRead(24, 2, 37), "in use by") > 0 Then
                        modMiscFunc.status_update "", "", "", "This macro is not stuck, WY order is locked by either a person or one of Automation. It keeps retrying to access in for 30 secs before dumping it to exception list"
                        .callSleep 2000
                        .setCursor 4, 31
                        .EnterKey
                    ElseIf InStr(.scrRead(24, 2, 64), "deactivated") > 0 Then
                        aErrCode = "ERR_WY has been deactivated"
                        .FunctionKey_12
                        Exit For
                    Else
                        Exit For
                    End If
                Next i
                If InStr(.scrRead(1, 23, 56), "Service Order Entry") > 0 Then
                    For i = 0 To 20
                        If InStr(.scrRead(5, 47, 66), "reserved") > 0 Then
                            modMiscFunc.status_update "", "", "", "It seems that UR order hasn`t been closed yet; Awaiting it to be posted"
                            .setCursor 18, 64
                            .SendKeyString "10"
                            .EnterKey
                            .callSleep 1500
                            .FunctionKey_01
                        Else
                            Exit For
                        End If
                    Next i
                    If Left(.scrRead(7, 8, 9), 1) = "5" Or .scrRead(7, 8, 9) = "A2" Then
                        If .scrRead(7, 39, 42) = "OPEN" Then
                            .setCursor 18, 64
                            For i = 0 To 51
                                If i = 51 Then
                                    aErrCode = "ERR_L/C_NOT_ATTACHED_ON_WY"
                                    Do While .scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE"
                                        .FunctionKey_03
                                    Loop
                                    Exit For
                                End If
                                If InStr(.scrRead(5, 47, 66), "reserved") > 0 Then
                                    .callSleep 500
                                    .setCursor 18, 64
                                    .FunctionKey_12
                                    Call EnterIntoExistingSO(SoNo)
                                Else
                                    Exit For
                                End If
                            Next i
                            If aErrCode = "" Then
                                .callSleep 500
                                .ConnectPS
                                .setCursor 18, 64
                                .SendKeyString "01"
                                .EnterKey
                                If .scrRead(6, 6, 76) = "Address Shown Used As Default For Line Cards Entered Through This Order" Then
                                    .FunctionKey_01
                                    If DASHBOARD.chkNotApplyBauCode = True Then ChangetoBTCode
                                    .ConnectPS
                                    .setCursor 18, 64
                                    .SendKeyString "07": .EnterKey
                                    .FunctionKey_12
                                    For i = 0 To 1
                                        .setCursor 8, 47
                                        If i = 0 Then
                                            .FieldExitKey
                                        Else
                                            .SendKeyString aIntrlve
                                            .EnterKey
                                            .FunctionKey_12
                                        End If
                                    Next
                                    .setCursor 18, 64
                                    .SendKeyString "33"
                                    .EnterKey
                                    If Trim(.scrRead(9, 12, 22)) <> "" Then
                                        .setCursor 9, 12
                                        .FieldExitKey
                                    End If
                                    If Trim(.scrRead(9, 53, 63)) <> "" Then
                                        For i = 0 To 1
                                            If i = 0 Then
                                                .setCursor 8, 53
                                            Else
                                                .setCursor 9, 53
                                            End If
                                            .FieldExitKey
                                        Next i
                                    End If
                                    .EnterKey
                                    .FunctionKey_01
                                    .setCursor 18, 64: .SendKeyString "10": .EnterKey
                                    .FunctionKey_08
                                    .setCursor 3, 2
                                    .SendKeyString "I1"
                                    .EnterKey
                                    .setCursor 3, 2
                                    .SendKeyString SETTING.Range("B13").value
                                    .EnterKey
                                    .FunctionKey_01
                                    temStr = .scrRead(7, 8, 9)
                                    If Not temStr = "A2" Or temStr = "MW" Or temStr = "GS" Or temStr = "92" Then
                                        .FunctionKey_09
                                        .setCursor 21, 52
                                        .SendKeyString "A2"
                                        .EnterKey
                                        .FunctionKey_12
                                    End If
                                    Do While .scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE"
                                        .FunctionKey_03
                                    Loop
                                End If
                            End If
                        Else
                            .setCursor 18, 64
                            .SendKeyString "17"
                            .EnterKey
                            If Trim(.scrRead(7, 29, 39)) = "" Then
                                aSONo = 0
                            Else
                                aSONo = Trim(.scrRead(7, 29, 39))
                            End If
                            .FunctionKey_01
                            If aErrCode = "" Then
                                aErrCode = "ERR_NOT_AN_OPEN_ORDER"
                            Else
                                aErrCode = aErrCode & vbNewLine & "ERR_NOT_AN_OPEN_ORDER"
                            End If
                            Do While .scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE"
                                .FunctionKey_03
                            Loop
                        End If
                    Else
                        For i = 0 To 1
                            .ConnectPS
                            .setCursor 18, 64
                            If i = 0 Then
                                Call ChangetoBTCode
                            End If
                            If i = 1 Then
                                .SendKeyString "17"
                                .EnterKey
                                temStr = Trim(.scrRead(7, 29, 39))
                                If IsNumeric(temStr) = True Then
                                    aSONo = Trim(.scrRead(7, 29, 39))
                                End If
                                .FunctionKey_01
                            End If
                        Next i
                        If aErrCode = "" Then
                            aErrCode = "ERR_WY_NOT_IN_5R"
                        Else
                            aErrCode = aErrCode & vbNewLine & "ERR_WY_NOT_IN_5R"
                        End If
                        .ConnectPS
                        BacktoChorusMenu
                    End If 'Left(.scrRead(7, 8, 9), 1) = "5" Or .scrRead(7, 8, 9) = "A2"
                End If 'instr(.scrRead(1,23,56),"Service Order Entry")>0
            End If
        End If
    End With
    Set icms = Nothing
End Sub

Function StagetoLineQuery() As Boolean
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String
    With icms
        If .ConnectPS = True Then
            If .onChorusMainMenu = True Then
                For i = 0 To 1
                    .setCursor 24, 58
                    If i = 0 Then
                        .SendKeyString "10"
                    Else
                        .SendKeyString "02"
                    End If
                    .EnterKey
                Next
                StagetoLineQuery = True
            ElseIf .scrRead(2, 33, 49) = "LINE CARD ENQUIRY" Then
                StagetoLineQuery = True
            Else
                StagetoLineQuery = False
            End If
        End If
    End With
    Set icms = Nothing
End Function

Function checkLWCusingASID(ByVal ucllASID As Long) As String
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String
    With icms
        If .ConnectPS = True Then
            If .scrRead(2, 33, 49) = "LINE CARD ENQUIRY" Then
                .FunctionKey_10
                For i = 0 To 1
                    .setCursor 4, 28
                    If i = 0 Then
                        .FieldExitKey
                    Else
                        .SendKeyString CStr(ucllASID)
                        .EnterKey
                    End If
                Next
                .setCursor 17, 60
                .SendKeyString "3"
                .EnterKey
                temStr = .scrRead(8, 53, 57)
                temStr = Trim(Replace(temStr, "-M", ""))
                checkLWCusingASID = temStr
                .FunctionKey_01
            End If
        End If
    End With
    Set icms = Nothing
End Function

Sub StagebackMainMenufromLineQuery()
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String
    With icms
        If .ConnectPS = True Then
            If .scrRead(2, 33, 49) = "LINE CARD ENQUIRY" Then
                .FunctionKey_03
                .setCursor 24, 58
                .SendKeyString "98"
                .EnterKey
            End If
        End If
    End With
    Set icms = Nothing
End Sub

Function CreateTLorder(InPotsLine As Long) As Long
    Dim icms As New clsICMS, i As Integer, temStr As String
    Dim newCustNo As Long
    Dim WinAPI As New clsWinAPI
    With icms
        If .ConnectPS = True Then
            If .onChorusMainMenu = True Then
                Call StagetoWWC
                Call EnterAcctIDinWWC
                .ConnectPS
                newCustNo = Trim(.scrRead(7, 70, 78))
                For i = 0 To 2
                    .FunctionKey_05
                Next
                Call EnterASIDinWWC
                .ConnectPS
                .EnterKey
                temStr = .scrRead(24, 2, 70)
                If InStr(temStr, "Line number not found in Subscriber") > 0 Then
                    aErrCode = "POTS line not Found"
                    .FunctionKey_03
                    Set icms = Nothing
                    Set WinAPI = Nothing
                    Exit Function
                End If
                temStr = .scrRead(7, 70, 78)
                If temStr = newCustNo Then
                    .FunctionKey_05
                    .FunctionKey_05
                    .FunctionKey_12
                    Set icms = Nothing
                    Set WinAPI = Nothing
                    Exit Function
                End If
                .setCursor 21, 19
                .SendKeyString "04"
                .EnterKey
                .FunctionKey_06
                .setCursor 3, 29
                For i = 0 To 2
                    .FieldExitKey
                Next
                .setCursor 3, 29
                .SendKeyString "TL"
                .setCursor 4, 29
                .SendKeyString Right(aTNo, Len(aTNo) - 1)
                .EnterKey
                .setCursor 7, 28
                .FieldExitKey
                .setCursor 7, 28
                .SendKeyString CStr(newCustNo)
                .EnterKey
                .setCursor 19, 73
                .SendKeyString Format(Now(), "ddmmyy")
                For i = 0 To 3
                    .FunctionKey_06
                Next
                .setCursor 3, 2
                .SendKeyString "Change Customer No. from " & temStr & " to " & newCustNo
                .FunctionKey_08
                .SendKeyString "I1"
                .EnterKey
                .setCursor 3, 2
                .SendKeyString "Request ID: " & aOOTId
                .FunctionKey_06
                .setCursor 18, 64
                .SendKeyString "17"
                .EnterKey
                aTNo = Trim(Replace(aTNo, "T", ""))
                .setCursor 13, 29
                .SendKeyString aTNo
                .setCursor 14, 29
                .SendKeyString aTNo
                .setCursor 19, 29
                .FunctionKey_08
                .setCursor 19, 72
                .FunctionKey_08
                .FunctionKey_23
                .FunctionKey_01
                Do While .scrRead(2, 23, 49) <> "WORK WITH CUSTOMER"
                    .FunctionKey_12
                Loop
                .setCursor 21, 19
                For i = 0 To 12
                    .FieldExitKey
                Next
                .setCursor 21, 19
                .SendKeyString "04"
                .EnterKey
                .FunctionKey_12
                .setCursor 4, 25
                For i = 0 To 10
                    .FieldExitKey
                Next
                Do While .scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE"
                    .FunctionKey_12
                Loop
            End If
        End If
    End With
    'Status Comments Needed
    Call status_update("", "", "", "The Macro is not stuck. Just await TL order to be posted")
    WinAPI.Sleeping 10000
    Set icms = Nothing
    Set WinAPI = Nothing
End Function

Sub ICMSHotKey()
    Dim icms As New clsICMS, i As Integer, temStr As String
    Dim icmsHD As Long, achild As Long, achild2 As Long
    Dim RUMBAFramerOCXWnd As Long, RUMBAFramerDocWnd As Long, WdPageFrame As Long
    Dim WdEngineWindow As Long, AfxOleControl As Long, AfxWnd As Long
    Dim WinAPI As New clsWinAPI
    With icms
        .ConnectPS
        'temStr = .scrRead(
        icmsHD = WinAPI.getObjectHwnd(vbNullString, "icms - Micro Focus Rumba")
        achild = WinAPI.getChildObjectHwnd(icmsHD, 0&, vbNullString, vbNullString)
        RUMBAFramerOCXWnd = WinAPI.getChildObjectHwnd(achild, 0&, "RUMBAFramerOCXWnd", vbNullString)
        RUMBAFramerDocWnd = WinAPI.getChildObjectHwnd(RUMBAFramerOCXWnd, 0&, "RUMBAFramerDocWnd", vbNullString)
        WdPageFrame = WinAPI.getChildObjectHwnd(RUMBAFramerDocWnd, 0&, "WdPageFrame", vbNullString)
        WdEngineWindow = WinAPI.getChildObjectHwnd(WdPageFrame, 0&, "WdEngineWindow", vbNullString)
        AfxOleControl = WinAPI.getChildObjectHwnd(WdEngineWindow, 0&, vbNullString, vbNullString)
        AfxWnd = WinAPI.getChildObjectHwnd(AfxOleControl, 0&, vbNullString, vbNullString)
        Call WinAPI.PressVK_ESC(AfxWnd)
        '.setCursor 20, 55
        '.SendKeyString "3"
        WinAPI.Sleeping 100
        .FunctionKey_06
        WinAPI.Sleeping 250
    End With
    Set icms = Nothing
    Set WinAPI = Nothing
End Sub


Function ScheduleHandler(ByVal scheduleCode As String) As String
'aStartDate
    Call ICMSHotKey
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String, row As Integer, tempArr() As String
    Dim month As String, nextMonth As String, pickedMonth As String, pickedDay As String, pickedAMorPM As String
    Dim aYear As String
    Dim MonthList As New Collection
    Dim amy1 As Integer, amy2 As Integer, pmy1 As Integer, pmy2 As Integer, offset As Integer, ISRCode As Integer
    Dim amVol As String, pmVol As String, chkNextMonth As Boolean
    
    If icms.ConnectPS = False Then Exit Function
    If icms.scrRead(1, 33, 48) = "Display Messages" Then
        icms.FunctionKey_13
        icms.FunctionKey_12
    End If
    If icms.scrRead(6, 28, 50) <> "CHORUS CUSTOMER SERVICE" Then
        MsgBox "Please stage back to '*** CHORUS CUSTOMER SERVICE ***' screen in ICMS, then click OK to continue", vbOKOnly, "UCLL CWO MACRO MESSAGE BOX"
    End If
    Call StagetoScheduleScrn
    If aStartDate = "" Then
        If aErrCode = "" Then
            aErrCode = "ERR_Failed to pick a valid RFS Date"
        Else
            aErrCode = aErrCode & vbNewLine & "ERR_Failed to pick a valid RFS Date"
        End If
        ScheduleHandler = "01011901AM"
        Set icms = Nothing
        Exit Function
    End If
    chkNextMonth = True
    aYear = Format(Now(), "yy")
    month = Left(Right(aStartDate, 4), 2)
    If month = "12" Then
        nextMonth = "01"
        aYear = CInt(aYear + 1)
    Else
        nextMonth = Format(CInt(month) + 1, "00")
    End If
    If Len(aStartDate) = 6 Then
        temStr = Left(aStartDate, 2)
    ElseIf Len(aStartDate) = 5 Then
        temStr = Left(aStartDate, 1)
    End If
    With icms
        If .ConnectPS = True Then
            If Not .scrRead(2, 24, 52) = "Planning Schedule Maintenance" Then
                Set icms = Nothing
                Set MonthList = Nothing
                If aErrCode = "" Then
                    aErrCode = "ERR_Failed to pick a valid RFS Date"
                Else
                    aErrCode = aErrCode & vbNewLine & "ERR_Failed to pick a valid RFS Date"
                End If
                ScheduleHandler = "01011901AM"
                Exit Function
            End If
            Do While chkNextMonth = True
                If .scrRead(2, 24, 52) = "Planning Schedule Maintenance" Then
                    .setCursor 5, 31
                    .SendKeyString month
                    .setCursor 6, 31
                    .SendKeyString scheduleCode
                    .EnterKey
                    If .scrRead(24, 2, 30) = "Schedule code  does not exist" Then
                        aErrCode = "ERR_Invalid Schedule Code"
                    End If
                    For row = 13 To 20
                        amy1 = 5: amy2 = 10
                        pmy1 = 13: pmy2 = 18
                        offset = 20
                        For i = 1 To 4
                            If (row = 20) And (i = 4) Then Exit For
                            amVol = Trim(.scrRead(row, amy1, amy2)): If InStr(amVol, "-") > 0 Or amVol = "" Then amVol = 0
                            pmVol = Trim(.scrRead(row, pmy1, pmy2)): If InStr(pmVol, "-") > 0 Or pmVol = "" Then pmVol = 0
                            MonthList.Add (amVol & ";" & pmVol)
                            amy1 = amy1 + offset: amy2 = amy2 + offset
                            pmy1 = pmy1 + offset: pmy2 = pmy2 + offset
                        Next
                    Next
                    pickedMonth = vbNullString: pickedDay = vbNullString: pickedAMorPM = vbNullString
                    For i = CInt(temStr) To MonthList.Count
                        If pickedDay <> vbNullString Then
                            chkNextMonth = False
                            Exit For
                        End If
                        tempArr = Split(MonthList.item(i), ";")
                        amVol = tempArr(0): pmVol = tempArr(1)
                        If amVol > 0 Or pmVol > 0 Then
                            pickedMonth = .scrRead(5, 31, 32)
                            pickedDay = Format(i, "00")
                            If amVol > 0 Then pickedAMorPM = "AM"
                            If pickedAMorPM = "" Then
                                If pmVol > 0 Then pickedAMorPM = "PM"
                            End If
                        End If
                    Next i
                    If pickedDay = vbNullString Then
                        If month = nextMonth Then Exit Do
                        month = nextMonth
                        chkNextMonth = True
                        Set MonthList = Nothing
                        Set MonthList = New Collection
                        temStr = 1
                        .FunctionKey_12
                        .setCursor 5, 31
                        For i = 0 To 2
                            .FieldExitKey
                        Next
                    End If
                End If
            Loop 'Do While chkNextMonth = True
            aStartDate = pickedDay & pickedMonth & aYear
            aStartTime = pickedAMorPM
            ScheduleHandler = pickedDay & pickedMonth & aYear & pickedAMorPM
        End If
    End With
    Call BacktoChorusMenu
    Call ICMSHotKey
    Set icms = Nothing
    Set MonthList = Nothing
End Function

Function GetICMSSOMfromAnOrder() As Long
    Dim icms As New clsICMS, i As Integer, i2 As Integer, temStr As String
    With icms
        If .ConnectPS = True Then
            If .scrRead(1, 23, 56) = "Service Order Entry / Enquiry Menu" Then
                .setCursor 18, 64
                .SendKeyString "17"
                .EnterKey
                temStr = Trim(.scrRead(7, 29, 40))
                If temStr = "" Then
                    GetICMSSOMfromAnOrder = 0
                Else
                    GetICMSSOMfromAnOrder = CLng(temStr)
                End If
                .FunctionKey_01
            End If
        End If
    End With
    Call BacktoChorusMenu
    Set icms = Nothing
End Function
