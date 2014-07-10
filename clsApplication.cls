VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const MYDBASETABLE = "\VBSampleApp.mda"

Private mvarModuleID As String 'local copy
Private mvarModuleDescr As String 'local copy
Private InfoTest As clsInfoTest

Public Property Get QuasiAddIn() As Boolean
   QuasiAddIn = True
End Property

Public Sub Initialize()
   Call Initialize2(CreateObject("Quasi97.application"))
End Sub

Public Sub Initialize2(ByRef qs As Object)

On Error GoTo errorhandler
   Set qst = qs
   Set Test1 = New VbSampleTest1
   Set Test2 = New VbSampleTest2
   
   Call qst.QuasiParameters.RegisterTest(Test1, Test1.TestID)
   Call qst.QuasiParameters.RegisterTest(Test2, Test2.TestID)
   Set InfoTest = New clsInfoTest
   Call InfoTest.Initialize(qst.StaticTests)
   'Call qst.CurPreampParameters.colHdTests.Add(InfoTest, InfoTest.ID)
   
'multi-setup example
    Set Dbptr = OpenDatabase(qst.QuasiParameters.SetupTmpFileName)
    Call gDbase.SynchronizeDatabasesByRef(App.Path & MYDBASETABLE, Dbptr)
    Call qst.QuasiParameters.RegisterTestClass("Sample Test 3", "vbAppSample.vbSampleTest3", "")
    
errorhandler:
     Select Case errorhandler("VbSampleApp:Initialize2")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Public Property Let ModuleDescr(ByVal vData As String)
    mvarModuleDescr = vData
End Property

Public Property Get ModuleDescr() As String
    ModuleDescr = mvarModuleDescr
End Property

Public Property Let ModuleID(ByVal vData As String)
    mvarModuleID = vData
End Property

Public Property Get ModuleID() As String
    ModuleID = mvarModuleID
End Property

Private Sub Class_Initialize()
    mvarModuleID = "vbSampleApp"
    mvarModuleDescr = "vb Sample App Tests"
End Sub

Private Sub Class_Terminate()
    'Workspaces(0).Close
    On Error GoTo errorhandler
    Call qst.QuasiParameters.UnRegisterTest(Test1, Test1.TestID)
    Set Test1 = Nothing
    Call qst.QuasiParameters.UnRegisterTest(Test2, Test2.TestID)
    Set Test2 = Nothing
    Call qst.QuasiParameters.UnregisterTestClass("Sample Test 3")
    
    Set Dbptr = Nothing
    Set qst = Nothing
    Err.Clear
    
errorhandler:
     Select Case errorhandler("VbSampleApp:Terminate")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

