Attribute VB_Name = "modMain"
Option Explicit

#If FLAG_DEBUG = 1 Then
   Public qst As Quasi97.Quasi97_Application
#Else
   Public qst As Object
#End If

Global Test1 As VbSampleTest1
Global Test2 As VbSampleTest2
Global gDbase As New ISIServMod.DataBaseFunc
Global Dbptr As Database

Global Const MAXSTATCOLUMN = 7

Global Const CONFIGHGA = 0
Global Const CONFIGHSA = 1
Global Const CONFIG8xHGA = 2
Global Const CONFIG2XHGA = 3
Global Const CONFIG2XBAR = 4
Global Const CONFIGBAR = 5
Global Const CONFIGTRAY = 6
Global Const CONFIG32xHGA = 7
Global Const CONFIG2XBARG3 = 8
Global Const CONFIG2xWHR = 9

Public Enum StressMode
    strStoreParams = 1
    strReStoreParams = 2
    strRunStress = 3
End Enum

Sub Main()

End Sub

Public Sub RegisterResult_(ResultID$, ByRef colresultnames As Collection)
    Dim rslt As Object
    Set rslt = qst.CreateCOMObj("Quasi97.ResultName")
    rslt.ResultID = ResultID
    colresultnames.Add rslt, ResultID
End Sub


Function InitResult() As Object
Dim rslt As Object
   Set rslt = qst.CreateCOMObj("Quasi97.Result")
   Set rslt.Results = New XArrayDB
   Set rslt.data = New XArrayDB
   Set rslt.Parameters = New XArrayDB
   Set rslt.Grades = New XArrayDB
   
   Call rslt.Results.ReDim(0, -1, 0, 0)
   Call rslt.data.ReDim(0, -1, 0, 0)
   Call rslt.Parameters.ReDim(0, -1, 0, 0)
   Call rslt.Grades.ReDim(0, -1, 0, 0)
   Set InitResult = rslt
End Function

Function errorhandler%(FuncName$)
    Dim Stat%
        
    If Err = -1 Then
        errorhandler = vbAbort
        Exit Function
    End If
    If Err = 401 Then 'trying to show non-modal form when modal is displayed
      errorhandler = vbIgnore
      Exit Function
    End If
    
    If Err <> 0 Then
        errorhandler = MsgBox(CStr(Err) + " [" & FuncName & "] : " + Error, vbAbortRetryIgnore, Err.Source)
    'Else
    '    errorhandler = vbAbort
    End If
End Function

Sub ResizeControls(frm As Form)
    On Error GoTo errorhandler
    
   frm.Left = 0
   frm.Top = 0
   frm.Width = qst.GetMainMod.ChildAreaWidth
   frm.Height = qst.GetMainMod.ChildAreaHeight

errorhandler:
     Select Case errorhandler("ResizeControls")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select

End Sub

