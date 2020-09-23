VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Main 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3122
   ClientLeft      =   42
   ClientTop       =   322
   ClientWidth     =   4550
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3122
   ScaleWidth      =   4550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   392
      Left            =   2016
      TabIndex        =   8
      Top             =   2646
      Width           =   1148
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   392
      Left            =   3276
      TabIndex        =   9
      Top             =   2646
      Width           =   1148
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify input file..."
      Height          =   392
      Left            =   126
      TabIndex        =   7
      Top             =   2646
      Width           =   1778
   End
   Begin VB.Frame Frame1 
      Height          =   2408
      Left            =   126
      TabIndex        =   10
      Top             =   90
      Width           =   4298
      Begin VB.TextBox txtReplace 
         Height          =   266
         Left            =   2268
         TabIndex        =   5
         Top             =   1640
         Width           =   1778
      End
      Begin VB.TextBox txtInput 
         Height          =   266
         Left            =   252
         TabIndex        =   0
         Top             =   378
         Width           =   3416
      End
      Begin VB.CommandButton cmdInputFile 
         Caption         =   "...."
         Height          =   336
         Left            =   3780
         TabIndex        =   1
         ToolTipText     =   "Click here to select the input file"
         Top             =   330
         Width           =   392
      End
      Begin VB.TextBox txtOutput 
         Height          =   266
         Left            =   252
         TabIndex        =   2
         Top             =   1008
         Width           =   3416
      End
      Begin VB.CheckBox chkEmbedded 
         Caption         =   "Replace embedded text"
         Height          =   266
         Left            =   252
         TabIndex        =   6
         Top             =   2016
         Width           =   2030
      End
      Begin VB.CommandButton Command2 
         Caption         =   "...."
         Height          =   336
         Left            =   3780
         TabIndex        =   3
         ToolTipText     =   "Click here to select the Output file"
         Top             =   970
         Width           =   392
      End
      Begin VB.TextBox txtSearch 
         Height          =   266
         Left            =   252
         TabIndex        =   4
         Top             =   1640
         Width           =   1778
      End
      Begin MSComctlLib.ProgressBar pgbStatus 
         Height          =   266
         Left            =   210
         TabIndex        =   11
         Top             =   2646
         Visible         =   0   'False
         Width           =   3920
         _ExtentX        =   6909
         _ExtentY        =   457
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog cmdialog 
         Left            =   2772
         Top             =   2016
         _ExtentX        =   813
         _ExtentY        =   813
         _Version        =   393216
      End
      Begin VB.Label lblStatusMessage 
         Caption         =   "0                                         50                                     100"
         Height          =   266
         Left            =   252
         TabIndex        =   16
         Top             =   2394
         Visible         =   0   'False
         Width           =   3794
      End
      Begin VB.Label Label4 
         Caption         =   "Replace with:"
         Height          =   266
         Left            =   2268
         TabIndex        =   15
         Top             =   1380
         Width           =   1400
      End
      Begin VB.Label Label1 
         Caption         =   "Input file:"
         Height          =   266
         Left            =   252
         TabIndex        =   14
         Top             =   150
         Width           =   1400
      End
      Begin VB.Label Label2 
         Caption         =   "Output file:"
         Height          =   266
         Left            =   252
         TabIndex        =   13
         Top             =   756
         Width           =   1400
      End
      Begin VB.Label Label3 
         Caption         =   "Search for:"
         Height          =   266
         Left            =   252
         TabIndex        =   12
         Top             =   1380
         Width           =   1400
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdInputFile_Click()
 txtInput.Text = LoadFile(Me.cmdialog, "C:\", "Text (*.txt)|*.txt|", "Select the input file")
End Sub

Private Sub cmdOk_Click()
 
 If CheckAllOk(txtInput, txtOutput, txtSearch, txtReplace) Then
   Call ActivatePanel
   pgbStatus.Value = 0
   lblStatusMessage.Visible = True
   pgbStatus.Visible = True
   Call SearchAndReplace(txtInput, txtOutput, txtSearch, txtReplace, chkEmbedded, pgbStatus)
   lblStatusMessage.Visible = False
   pgbStatus.Visible = False
 End If
 
End Sub

Private Sub cmdVerify_Click()
Dim Temp As Boolean
 
  Temp = VerifyFileOk(txtInput)
  If Temp Then MsgBox "Check complete! The file is OK.", vbInformation + vbOKOnly, APP_Name
  
End Sub

Private Sub Command2_Click()
 txtOutput.Text = LoadFile(Me.cmdialog, "C:\", "Text (*.txt)|*.txt|", "Select the Output file")
End Sub

Private Sub Form_Load()
 Me.Caption = APP_Name & " " & APP_Version
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer

 Response = MsgBox("Are you sure that you want to exit?", vbInformation + vbOKCancel, "Exit from SearchMaster")
 If Response = vbCancel Then
   Cancel = True
 End If
  
End Sub

Sub ActivatePanel()
  Frame1.Height = 3038
  Me.Height = 4130
  cmdVerify.Top = 3276
  cmdOk.Top = 3276
  cmdExit.Top = 3276
End Sub

