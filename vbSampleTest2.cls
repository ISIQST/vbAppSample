VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbSampleTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' required
Public ContainsGraph As Integer
Public InfoColumns As New Collection
Public colresultnames As New Collection
Public colResults As New Collection
Private rslt As Object 'Quasi97.Result
Public RefreshLogHeader As Boolean      'Required for data logging
Private mvarTestID As String
Private mvarTestForm As frmVBSampleTest2
Private lInfo() As Variant
Public FeatureVector As Integer
Private mvarBias As Single

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
    mvarBias = vData
End Property

Public Property Get Bias() As Single
    Bias = mvarBias
End Property

Public Property Set TestForm(ByVal vData As Object)
    Set mvarTestForm = vData
End Property

Public Property Get TestForm() As Object
    If mvarTestForm Is Nothing Then Set mvarTestForm = New frmVBSampleTest1
    If mvarTestForm.OwnerTest Is Nothing Then Set mvarTestForm.OwnerTest = Me
    Set TestForm = mvarTestForm
End Property

Public Property Let TestID(ByVal vData As String)
    mvarTestID = vData
End Property

Public Property Get TestID() As String
    TestID = mvarTestID
End Property

Public Sub RunTest()
    Dim SaveBias!, Res!
    Dim Bias!, BiasFrom!, BiasTo!, BiasInc!
    Dim i%, TotPts%
       
    On Error GoTo errorhandler
    ReDim lInfo(0 To 0, 0 To InfoColumns.Count - 1)
    
    Set rslt = InitResult
    rslt.Results.ReDim 0, rslt.Results.UpperBound(1), 0, 8
    If colResults.Count > 0 Then
        colResults.Add rslt, , 1
    Else
        colResults.Add rslt
    End If
    SaveBias = qst.SystemParameters.DefaultReadBiasCh1
    BiasFrom = Val(mvarTestForm.txtBiasFrom)
    BiasTo = Val(mvarTestForm.txtBiasTo)
    BiasInc = Val(mvarTestForm.txtBiasInc)
    
    TotPts = (BiasTo - BiasFrom) / BiasInc
    mvarTestForm.grdResult.Rows = TotPts + 1
    
    qst.QstStatus.Message = "Measuring Resistance"
    For i = 0 To TotPts - 1
        Bias = BiasFrom + i * BiasInc
        qst.SystemParameters.DefaultReadBiasCh1 = Bias
        Res = qst.QSTHardware.HReadRes(0, 10)
        mvarTestForm.grdResult.Row = i + 1
        mvarTestForm.grdResult.Col = 0
        mvarTestForm.grdResult.Text = Bias
        mvarTestForm.grdResult.Col = 1
        mvarTestForm.grdResult.Text = Res
        Call rslt.AddResult("Bias (mA)", Format(Bias, "##0.0#"), 1)
        Call rslt.AddResult("Resistance (Ohm)", Format(Res, "##0.0#"), 1)
     Next
     
    Call rslt.AddInfo(Me, 0)
    If Not qst.QuasiParameters.ProductionMode Then Call qst.QuasiParameters.LogResults(Me, 1)
     
errorhandler:
   Select Case errorhandler("vbSampleTest2:RunTest")
      Case vbRetry
          Resume
      Case vbIgnore
          Resume Next
   End Select
       
    qst.SystemParameters.DefaultReadBiasCh1 = SaveBias
    qst.QstStatus.Clear
        
End Sub

Public Sub RestoreParameters()

End Sub

Public Sub StoreParameters()

End Sub

Private Sub Class_Initialize()
       ContainsGraph = 2
    mvarTestID = "VbSampleTest2"
    Set mvarTestForm = New frmVBSampleTest2
    Set InfoColumns = New Collection
    Call qst.QuasiParameters.InitInfoCollection(InfoColumns)
    Call RegisterResult_("Resistance (Ohm)", colresultnames)
End Sub

Public Sub DisplayResults(rslt As Result)

End Sub

Public Function InitChartTransfer(Optional sizex As Long = 0, Optional sizey As Long = 0, Optional filename As String = "") As Boolean
   If filename = "" Then
      TestForm.grfData.EditCopy
   Else     'need to save the image of the chart to the disk under "filename". extension is not provided
      
   End If
   InitChartTransfer = True
End Function

Public Function EndChartTransfer() As Boolean
   Clipboard.Clear
   EndChartTransfer = True
End Function
