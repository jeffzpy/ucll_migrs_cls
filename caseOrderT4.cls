VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseOrderT4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim aOrder As clsOrderInterface

Private Sub Class_Initialize()
    Set aOrder = New clsOrderInterface
End Sub
Private Sub Class_Terminate()
    Set aOrder = Nothing
End Sub

Public Property Get AccountNo() As Long
    AccountNo = aOrder.AccountNo
End Property

Public Property Let AccountNo(ByVal AccountNo As Long)
    aOrder.AccountNo = AccountNo
End Property

Public Property Get ootid() As Long
    ootid = aOrder.ootid
End Property

Public Property Let ootid(ByVal ootid As Long)
    aOrder.ootid = ootid
End Property

Public Property Get rspRef() As String
    rspRef = aOrder.rspRef
End Property

Public Property Let rspRef(ByVal rspRef As String)
    aOrder.rspRef = rspRef
End Property

Public Property Let LspRef(ByVal LspRef As String)
    aOrder.LspRef = LspRef
End Property

Public Property Let LspName(ByVal LspName As String)
    aOrder.LspName = LspName
End Property

Public Property Let GainName(ByVal GainName As String)
    aOrder.GainName = GainName
End Property

Public Property Get ExistingAsid() As Long
    ExistingAsid = aOrder.ExistingAsid
End Property

Public Property Let ExistingAsid(ByVal ExistingAsid As Long)
    aOrder.ExistingAsid = ExistingAsid
End Property

Public Property Get ExistingUCLLAsid() As Long
    ExistingUCLLAsid = aOrder.ExistingUCLLAsid
End Property

Public Property Let ExistingUCLLAsid(ByVal ExistingUCLLAsid As Long)
    aOrder.ExistingUCLLAsid = ExistingUCLLAsid
End Property

Public Property Get StartDate() As String
    StartDate = aOrder.StartDate
End Property

Public Property Let StartDate(ByVal StartDate As String)
    aOrder.StartDate = StartDate
End Property

Public Property Get StartTime() As String
    StartTime = aOrder.StartTime
End Property

Public Property Let StartTime(ByVal StartTime As String)
    aOrder.StartTime = StartTime
End Property

Public Property Get OrderType() As String
    OrderType = aOrder.OrderType
End Property

Public Property Let OrderType(ByVal OrderType As String)
    aOrder.OrderType = OrderType
End Property

Public Property Get tNo() As String
    tNo = aOrder.tNo
End Property

Public Property Let tNo(ByVal tNo As String)
    aOrder.tNo = tNo
End Property

Public Property Let BTCode(ByVal BTCode As String)
    aOrder.BTCode = BTCode
End Property

Public Property Get ProvPlatform() As String
    ProvPlatform = aOrder.ProvPlatform
End Property

Public Property Let ProvPlatform(ByVal ProvPlatform As String)
    aOrder.ProvPlatform = ProvPlatform
End Property

Public Property Get exchange() As String
    exchange = aOrder.exchange
End Property

Public Property Let exchange(ByVal exchange As String)
    aOrder.exchange = exchange
End Property

Public Property Get samID() As Long
    samID = aOrder.samID
End Property

Public Property Let samID(ByVal samID As Long)
    aOrder.samID = samID
End Property

Public Property Get SpotArr() As Variant
    SpotArr = aOrder.SpotArr
End Property

Public Property Let SpotArr(ByVal SpotArr As Variant)
    aOrder.SpotArr = SpotArr
End Property

Public Property Get Intrlve() As String
    Intrlve = aOrder.Intrlve
End Property

Public Property Let Intrlve(ByVal Intrlve As String)
    aOrder.Intrlve = Intrlve
End Property

Public Property Get CustName() As String
    CustName = aOrder.CustName
End Property

Public Property Let CustName(ByVal CustName As String)
    aOrder.CustName = CustName
End Property

Public Property Get CustContactNo() As String
    CustContactNo = aOrder.CustContactNo
End Property

Public Property Let CustContactNo(ByVal CustContactNo As String)
    aOrder.CustContactNo = CustContactNo
End Property

Public Property Get CustContactNo2() As String
    CustContactNo2 = aOrder.CustContactNo2
End Property

Public Property Let CustContactNo2(ByVal CustContactNo2 As String)
    aOrder.CustContactNo2 = CustContactNo2
End Property

Public Property Get EubaPlan() As String
    EubaPlan = aOrder.EubaPlan
End Property

Public Property Let EubaPlan(ByVal EubaPlan As String)
    aOrder.EubaPlan = EubaPlan
End Property

Public Property Get ErrCode() As String
    ErrCode = aOrder.ErrCode
End Property

Public Property Let ErrCode(ByVal ErrCode As String)
    aOrder.ErrCode = ErrCode
End Property

Public Property Get SoNo() As Long
    SoNo = aOrder.SoNo
End Property

Public Property Let SoNo(ByVal SoNo As Long)
    aOrder.SoNo = SoNo
End Property

Public Property Get CreatedOrderDetails() As String
    CreatedOrderDetails = aOrder.CreatedOrderDetails
End Property

Public Property Let CreatedOrderDetails(ByVal CreatedOrderDetails As String)
    aOrder.CreatedOrderDetails = CreatedOrderDetails
End Property

Public Property Get ExchangeMismatch() As Boolean
    ExchangeMismatch = aOrder.ExchMismatch
End Property

Public Function CreateOrderProcedure() As String
    'Call Class_Initialize
    With aOrder
        .CreateTLorder aOrder.ExistingAsid
        If .ErrCode <> "" Then
            CreateOrderProcedure = .ErrCode
            GoTo finish
        End If
        If .ErrCode <> "" Then
            Exit Function
        End If
        .BBorPOTS = "POTS"
        .OrderType = "T4"
        .ServiceType = "RI"
        .ProvPlatform = "NEAX"
        .StagetoWWC
        .EnterASIDinWWC
        .StagetoOrderEnter
        .StagetoInitialSOScrn
        .Fill_ISOE_Screen
        .Fill_SAM_Screen "X"
        .Fill_Schedule_Screen
        .Fill_SPOTS_Screen
        .Fill_Memo_Screen
        If .ExchMismatch = True Then
            .StageanOrdertoXL
        End If
        If DASHBOARD.chkNotApplyBauCode = True Then
            .ChangetoBTCode
        End If
        .BacktoChorusMenu
        CreateOrderProcedure = .CreatedOrderDetails
        Application.Wait (Now + TimeValue("0:00:01"))
    End With
finish:
    'Call Class_Terminate
End Function
