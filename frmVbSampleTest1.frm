VERSION 5.00
Begin VB.Form frmVBSampleTest1 
   Caption         =   "VB Samle Test 1"
   ClientHeight    =   6810
   ClientLeft      =   1890
   ClientTop       =   1770
   ClientWidth     =   11055
   Icon            =   "frmVbSampleTest1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleMode       =   0  'User
   ScaleWidth      =   12194.83
   Begin VB.TextBox txtBias 
      Height          =   315
      Left            =   7440
      TabIndex        =   1
      Text            =   "5"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label MeasRes 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   7440
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblMeasRes 
      Alignment       =   1  'Right Justify
      Caption         =   "Measured Res:"
      Height          =   315
      Left            =   5820
      TabIndex        =   2
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label lblBias 
      Alignment       =   1  'Right Justify
      Caption         =   "Bias (mA)"
      Height          =   315
      Left            =   6180
      TabIndex        =   0
      Top             =   360
      Width           =   1035
   End
End
Attribute VB_Name = "frmVBSampleTest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OwnerTest As VbSampleTest1
Private Const ErrModId As String = "frmVBSampleTest1: "

Private Sub Form_Activate()
   If Not (qst.CurTestParameters Is OwnerTest) Then Set qst.CurTestParameters = OwnerTest
End Sub

Private Sub Form_Load()

   On Error GoTo errorhandler
   
   Call ResizeControls(Me)
   
errorhandler:
     Select Case errorhandler(ErrModId & "Load")
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call CloseFormSub
End Sub

Public Sub CloseForm()
   Call CloseFormSub
   Unload Me
End Sub

Private Sub CloseFormSub()
   If Not OwnerTest Is Nothing Then
      Call OwnerTest.StoreParameters
      If (qst.CurTestParameters Is OwnerTest) Then
          Set qst.CurTestParameters = Nothing
      Else
          Call qst.RemoveFromCurTest(OwnerTest)
      End If
      Set OwnerTest = Nothing
   End If
End Sub
