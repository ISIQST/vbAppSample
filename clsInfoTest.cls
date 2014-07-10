VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------
'this test class is for Quasi97 Static Test. It does not require any parameters, but it can supply the data to the log file.
Option Explicit

Public ID As String

Public Sub Execute(test As Object)
Dim I As Integer
Dim custom1 As String
Dim Custom2 As String
On Error Resume Next

   For I = 1 To test.Cycles
      Custom2 = "09876" & Rnd(2)
      Call test.colResults.Item(1).AddResult("Info1", "12456" & Rnd(1), 1)
      Call test.colResults.Item(1).AddResult("Info2", Custom2, 1)
   Next I
End Sub

Public Sub Initialize(test As Object)
    Call test.RegisterResult("Info1")
    Call test.RegisterResult("Info2")
End Sub

Private Sub Class_Initialize()
    ID = "SliderInfo"
End Sub
'---------------------------------
