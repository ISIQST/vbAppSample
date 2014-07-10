VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbSampleTest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' required
Public ContainsGraph As Boolean
Public colresultnames As New Collection
Public colResults As New Collection
Public InfoColumns As New Collection
Private rslt As Object
Public RefreshLogHeader As Boolean          'Required for logging
Public FeatureVector As Integer
Private mvarTestID As String
Private mvarTestForm As frmVBSampleTest1
Private mvarBias As Single
Private lInfo() As Variant

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
    Dim SaveBias!
       
    On Error GoTo errorhandler
    ReDim lInfo(0 To 0, 0 To InfoColumns.Count - 1)
    
    Set rslt = InitResult
    rslt.Results.ReDim 0, rslt.Results.UpperBound(1), 0, 8
    If colResults.Count > 0 Then
        colResults.Add rslt, , 1
    Else
        colResults.Add rslt
    End If
    SaveBias = qst.QSTHardware.GetBias(0)
    Call qst.QSTHardware.SetBias(0, Val(mvarTestForm.txtBias))
    qst.QstStatus.Message = "Measuring Resistance"
    mvarTestForm.MeasRes = qst.QSTHardware.HReadRes(0, 10)

    Call rslt.AddResult("Resistance (Ohm)", Format(mvarTestForm.MeasRes, "##0.0#"), 1)
     
    Call rslt.AddInfo(Me, 0)
    If Not qst.QuasiParameters.ProductionMode Then Call qst.QuasiParameters.LogResults(Me, 1)
     
errorhandler:
         Select Case errorhandler("vbSampleTest1:RunTest")
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
       
    Call qst.QSTHardware.SetBias(0, SaveBias)
    qst.QstStatus.Clear
        
End Sub

Public Sub RestoreParameters()

End Sub

Public Sub StoreParameters()

End Sub

Private Sub Class_Initialize()
    mvarTestID = "VbSampleTest1"
    Set mvarTestForm = New frmVBSampleTest1
    Set InfoColumns = New Collection
    Call qst.QuasiParameters.InitInfoCollection(InfoColumns)
    Call RegisterResult_("Resistance (Ohm)", colresultnames)
End Sub

Public Sub DisplayResults(rslt As Object)

On Error GoTo errorhandler


errorhandler:
    Select Case errorhandler("ESDTest : DisplayResults")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub
