VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmVBSampleTest3 
   Caption         =   "VB Samle Test 3"
   ClientHeight    =   6720
   ClientLeft      =   1560
   ClientTop       =   1200
   ClientWidth     =   9030
   Icon            =   "frmVbSampleTest3.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleMode       =   0  'User
   ScaleWidth      =   9961.042
   Begin VB.CheckBox chkStress 
      Enabled         =   0   'False
      Height          =   195
      Left            =   7350
      TabIndex        =   9
      Top             =   2340
      Width           =   195
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3555
      Left            =   270
      OleObjectBlob   =   "frmVbSampleTest3.frx":030A
      TabIndex        =   8
      Top             =   3120
      Width           =   7065
   End
   Begin VB.CommandButton cmdStress 
      Caption         =   "Stress"
      Height          =   645
      Left            =   7140
      TabIndex        =   7
      Top             =   2160
      Width           =   1725
   End
   Begin MSFlexGridLib.MSFlexGrid grdResult 
      Height          =   2355
      Left            =   300
      TabIndex        =   6
      Top             =   690
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
Attribute VB_Name = "frmVBSampleTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OwnerTest As VbSampleTest3
Private Const ErrModId As String = "frmVBSampleTest3: "

Private Sub cmdStress_Click()
   OwnerTest.Stress.ShowForm
   chkStress.Value = IIf(OwnerTest.Stress.StressList.Count(1) > 0, 1, 0)
End Sub

Private Sub Form_Activate()
   Me.Caption = OwnerTest.TestID & "." & OwnerTest.UserName
   chkStress.Value = IIf(OwnerTest.Stress.StressList.Count(1) > 0, 1, 0)
End Sub

Private Sub Form_Load()

    On Error GoTo errorhandler
   
   Call ResizeControls(Me)
   'Store handle to this form's window

   txtBiasFrom = OwnerTest.BiasFrom
   txtBiasTo = OwnerTest.BiasTo
   txtBiasInc = OwnerTest.BiasInc
   Call OwnerTest.RestoreParameters
   'Call procedure to begin capturing messages for this window
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

Private Sub txtBiasFrom_LostFocus()
   OwnerTest.BiasFrom = Val(txtBiasFrom)
End Sub

Private Sub txtBiasInc_LostFocus()
   OwnerTest.BiasInc = Val(txtBiasInc)
End Sub

Private Sub txtBiasTo_LostFocus()
   OwnerTest.BiasTo = Val(txtBiasTo)
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
