VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBSampleTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private testseq As Object
Private WithEvents TestSeq As Quasi97.TestSequencer
Attribute TestSeq.VB_VarHelpID = -1

Private Sub Class_Initialize()
   Set TestSeq = qst.TestSequencer
End Sub

Private Sub TestSeq_RunTestTerminate()
   Dim i%, n%, x%
   Dim xarr As Object
   For i = 1 To TestSeq.colResults.Count
      Set xarr = TestSeq.colResults.Item(i).Grades
      For n = 0 To xarr.UpperBound(1)
         For x = 0 To xarr.UpperBound(2)
            Debug.Print xarr.Value(n, x)
         Next
      Next
   Next
End Sub
