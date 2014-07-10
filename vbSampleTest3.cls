VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbSampleTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Setup As Integer
Public UserName As String
Private Dbptr As Database

' required
Public ContainsGraph As Integer
Public colresultnames As New Collection
Public colResults As New Collection
Public RefreshLogHeader As Boolean      'Required for data logging
Public DataColumns As New Collection
Public InfoColumns As New Collection
Public ColAdaptiveParameters As New Collection
Public FeatureVector As Integer
' required
Private mvarTestID As String
Private mvarTestForm As frmVBSampleTest3

' belong to this test
Private mvarBias As Single
Private RSParameters As Recordset
Public BiasFrom As Single
Public BiasTo As Single
Public BiasInc As Single
Public Stress As Object ' Quasi97.clsStress
Private lInfo() As Variant
Dim clpbrdset As Boolean

Private GValidateOption As ValidateOptions
Enum ValidateOptions
   enormal = 0
   eQuiet = 1
   eSkip = 2
End Enum

Public Property Let ValidateOption(ByVal valoption As Integer)
   GValidateOption = valoption
End Property

Public Property Let Info(idx1 As Integer, idx2 As Integer, data As Variant)
   lInfo(idx1, idx2) = data
End Property

Public Property Get Info(idx1 As Integer, idx2 As Integer) As Variant
   Info = lInfo(idx1, idx2)
End Property

Public Property Get NumberOfSetupsRun() As Integer
    NumberOfSetupsRun = 1
End Property

Public Sub ClearResults()
    Dim i%
    
    For i = 1 To colResults.Count
        colResults.Remove 1
    Next
    Set colResults = Nothing

End Sub

Public Property Let Bias(ByVal vData As Single)
    If GValidateOption = enormal Then
      If Abs(vData) > 5 Then
         Call MsgBox("The bias is out of range. " & CStr(vData) & " > 5", vbOKOnly)
      End If
   ElseIf GValidateOption = eQuiet Then
      If Abs(vData) > 5 Then vData = 5 * (vData / Abs(vData))
   End If
   mvarBias = vData
End Property

Public Property Get Bias() As Single
    Bias = mvarBias
End Property

Public Property Set TestForm(ByVal vData As Object)
    Set mvarTestForm = vData
End Property

Public Property Get TestForm() As Object
    Set TestForm = mvarTestForm
    If mvarTestForm.OwnerTest Is Nothing Then Set mvarTestForm.OwnerTest = Me
End Property

Public Property Let TestID(ByVal vData As String)
    mvarTestID = vData
End Property

Public Property Get TestID() As String
    TestID = mvarTestID
End Property

Public Sub RunTest()
    Dim SaveBias!, Res!
    Dim Bias!
    Dim i%, TotPts%
   Dim rslt As Object
   
    On Error GoTo errorhandler
    
    Set rslt = InitResult
    rslt.Results.ReDim 0, rslt.Results.UpperBound(1), 0, 8
    If colResults.Count > 0 Then
'        info.Insert 1, 0
        colResults.Add rslt, , 1
    Else
        colResults.Add rslt
    End If
    ReDim lInfo(0 To 0, 0 To InfoColumns.Count - 1)
    
   Call qst.Normalization.GetAdaptParam(TestID, Setup, "", 0)
   Call Stress.RunStress(StressMode.strStoreParams)         'store existing state of parameters. The parameters will be restored to this state at the end of the test
   
   SaveBias = qst.SystemParameters.DefaultReadBiasCh1
    
    TotPts = (BiasTo - BiasFrom) / BiasInc
    TestForm.grdResult.Rows = TotPts + 1
    'mvarTestForm.grfData.ColumnCount = 1
    'mvarTestForm.grfData.RowCount = TotPts
    
    
   qst.QstStatus.Message = "Measuring Resistance"
   Call rslt.data.ReDim(0, TotPts - 1, 0, 1)
   
   Call rslt.AddPreConditioningParameters(Me.Stress)        'record stress options under parameters (for output to the user), to be able to refer to test conditions later
   Call Stress.RunStress(StressMode.strRunStress)           'now run the test. Here this is done only once per test run.
   
    For i = 0 To TotPts - 1
        Bias = BiasFrom + i * BiasInc
        qst.SystemParameters.DefaultReadBiasCh1 = Bias
        Res = qst.QSTHardware.HReadRes(0, 10)
        mvarTestForm.grdResult.Row = i + 1
        mvarTestForm.grdResult.Col = 0
        'mvarTestForm.grfData.Row = i + 1
        'mvarTestForm.grfData.Column = 1
        mvarTestForm.grdResult.Text = Bias
        'mvarTestForm.grfData.Data = Res
        mvarTestForm.grdResult.Col = 1
        mvarTestForm.grdResult.Text = Res
        
        rslt.data.Value(i, 0) = Bias
        rslt.data.Value(i, 1) = Res
        
     Next
   
   Call rslt.AddResult("Bias (mA)", Format(Bias, "##0.0#"), 1)
   Call rslt.AddResult("Resistance (Ohm)", Format(Res, "##0.0#"), 1)
   
   Call qst.GradingParameters.GradeTest(Me, 1)
        
   Call qst.Normalization.AddResultforRecord(TestID, Setup, "", rslt)
    Call rslt.AddInfo(Me, 0)
    If Not qst.QuasiParameters.ProductionMode Then Call qst.QuasiParameters.LogResults(Me, 1)
     
errorhandler:
         Select Case errorhandler("vbSampleTest2:RunTest")
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
       
    Call qst.QSTHardware.SetBias(1, SaveBias)
    qst.QstStatus.Clear
    Call Stress.RunStress(StressMode.strReStoreParams)         'need to restore all of the parameters to the state before the stress was executed
    
End Sub

Public Sub RestoreParameters()
Dim NeedToStore As Boolean

    If Not RSParameters Is Nothing Then
    
      If RSParameters.RecordCount = 0 Then
         RSParameters.AddNew
         RSParameters.Update
         RSParameters.MoveFirst
         NeedToStore = True
      End If
        BiasFrom = RSParameters.Fields("BiasFrom")
        BiasTo = RSParameters.Fields("BiasTo")
        BiasInc = RSParameters.Fields("BiasInc")
    End If
    
    Call Stress.RestoreParameters(TestID, Setup)
    If NeedToStore Then Call StoreParameters
End Sub

Public Sub StoreParameters()
    If Not RSParameters Is Nothing Then
        RSParameters.Edit
        RSParameters.Fields("BiasFrom") = BiasFrom
        RSParameters.Fields("BiasTo") = BiasTo
        RSParameters.Fields("BiasInc") = BiasInc
        RSParameters.Fields("Setup") = Setup
        RSParameters.Update
    End If
    Call Stress.StoreParametersNew(TestID, Setup)
End Sub

Private Sub Class_Initialize()
    ContainsGraph = 2
    mvarTestID = "Sample Test 3"
    
    Set Stress = qst.CreateCOMObj("Quasi97.clsStress")
    
    Set mvarTestForm = New frmVBSampleTest3
    DataColumns.Add "Bias"
    DataColumns.Add "Resistance"
    
    Set InfoColumns = New Collection
    Call qst.QuasiParameters.InitInfoCollection(InfoColumns)
    
    ColAdaptiveParameters.Add "Bias"
    ColAdaptiveParameters.Add "BiasTo"
    FeatureVector = FeatureVector Or &H10
    
    Call RegisterResult_("Resistance (Ohm)", colresultnames)
End Sub

Public Sub DisplayResults(rslt As Result)

End Sub

Private Sub Class_Terminate()
   If Not RSParameters Is Nothing Then Set RSParameters = Nothing
   Set Dbptr = Nothing
End Sub

Public Sub RemoveRecord()
Dim i As Integer
'should delete the record of itself in the database
   If Not (RSParameters Is Nothing) Then
      On Error Resume Next
      Call RSParameters.MoveFirst
      Call RSParameters.Delete
      Set RSParameters = Nothing
   End If
   If Not Stress Is Nothing Then Stress.DeleteStressFromDB
End Sub

Public Sub SetDBase(ByVal NewDBase$, Optional ByRef AllSetups As Object = Nothing)
'all setups should have one item, if user wants to add one item by default.
   On Error GoTo errorhandler
   
   If NewDBase = "" Then
      Set Dbptr = Nothing
   Else
      Set Dbptr = OpenDatabase(NewDBase)
      If Not AllSetups Is Nothing Then       'special case: Quasi97 did not have any setups of that type and it asks the test how many  are present in the database
         Set RSParameters = Dbptr.OpenRecordset("Select * from VBSampleTest3")
         If RSParameters.RecordCount > 0 Then
            RSParameters.MoveLast
            RSParameters.MoveFirst
            
            While Not RSParameters.EOF
               AllSetups.Add RSParameters.Fields("Setup").Value         'add each setup to the database
               RSParameters.MoveNext
            Wend
         End If
      End If
      Set RSParameters = Dbptr.OpenRecordset("Select * from VBSampleTest3 Where Setup = " & Setup)
   End If
   
   Stress.OwnerSetupNumber = Setup
   Stress.OwnerTestID = mvarTestID
   Call Stress.SetDBase(NewDBase, Nothing)
   
errorhandler:
     Select Case errorhandler("SetDbase")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Public Function InitChartTransfer() As Boolean
   'Set clpbrdobj = Clipboard.GetData()
   Call TestForm.MSChart1.EditCopy
   InitChartTransfer = True
   clpbrdset = True
End Function

Public Function EndChartTransfer() As Boolean
   If clpbrdset Then
      Clipboard.Clear
      EndChartTransfer = True
   End If
End Function
