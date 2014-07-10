VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmVBSampleTest2 
   Caption         =   "VB Samle Test 2"
   ClientHeight    =   5985
   ClientLeft      =   1560
   ClientTop       =   1200
   ClientWidth     =   11055
   Icon            =   "frmVbSampleTest2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleMode       =   0  'User
   ScaleWidth      =   12194.83
   Begin MSChart20Lib.MSChart grfData 
      Height          =   3480
      Left            =   0
      OleObjectBlob   =   "frmVbSampleTest2.frx":030A
      TabIndex        =   7
      Top             =   0
      Width           =   5865
   End
   Begin MSFlexGridLib.MSFlexGrid grdResult 
      Height          =   2355
      Left            =   300
      TabIndex        =   6
      Top             =   3900
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4154
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.TextBox txtBiasInc 
      Height          =   315
      Left            =   7920
      TabIndex        =   5
      Text            =   "1"
      Top             =   1500
      Width           =   795
   End
   Begin VB.TextBox txtBiasTo 
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Text            =   "6"
      Top             =   1020
      Width           =   795
   End
   Begin VB.TextBox txtBiasFrom 
      Height          =   315
      Left            =   7920
      TabIndex        =   3
      Text            =   "2"
      Top             =   540
      Width           =   795
   End
   Begin VB.Label lblBiasInc 
      Alignment       =   1  'Right Justify
      Caption         =   "Bias Inc:"
      Height          =   315
      Left            =   6600
      TabIndex        =   2
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label lblBiasTo 
      Alignment       =   1  'Right Justify
      Caption         =   "Bias To:"
      Height          =   315
      Left            =   6600
      TabIndex        =   1
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label lblBiasFrom 
      Alignment       =   1  'Right Justify
      Caption         =   "Bias From:"
      Height          =   315
      Left            =   6600
      TabIndex        =   0
      Top             =   540
      Width           =   1035
   End
End
Attribute VB_Name = "frmVBSampleTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OwnerTest As VbSampleTest2
Private Const ErrModId As String = "frmVBSampleTest2: "

Private Sub Form_Load()

    On Error GoTo errorhandler
   
   Call ResizeControls(Me)
   grdResult.Row = 0
   grdResult.Col = 0
   grdResult.Text = "Bias (mA)"
   grdResult.Col = 1
   grdResult.Text = "Res (Ohm)"
   
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
