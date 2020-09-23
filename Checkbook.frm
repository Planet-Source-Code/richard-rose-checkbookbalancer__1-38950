VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Checkbook"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8595
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCurrentBalance 
      DataSource      =   "dtaCheckbook"
      Height          =   495
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtDescription 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   3800
      Width           =   2295
   End
   Begin VB.TextBox txtCheckNumber 
      Height          =   495
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtBalance 
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5000
      Width           =   1935
   End
   Begin VB.TextBox txtAmount 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2800
      Width           =   1815
   End
   Begin VB.Frame frmTransaction 
      BackColor       =   &H00808080&
      Caption         =   "Transaction"
      Height          =   1815
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton optTransaction 
         BackColor       =   &H00808080&
         Caption         =   "Withdrawal"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optTransaction 
         BackColor       =   &H00808080&
         Caption         =   "Deposit"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optTransaction 
         BackColor       =   &H00808080&
         Caption         =   "Check Card"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optTransaction 
         BackColor       =   &H00808080&
         Caption         =   "Check"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label lblCurrentBalance 
      BackColor       =   &H00808080&
      Caption         =   "Current Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00808080&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   3800
      Width           =   2055
   End
   Begin VB.Label lblCheckNumber 
      BackColor       =   &H00808080&
      Caption         =   "Check Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00808080&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00808080&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2800
      Width           =   1215
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00808080&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3120
      TabIndex        =   10
      Top             =   5000
      Width           =   1755
   End
   Begin VB.Menu mnuCompute 
      Caption         =   "&Compute"
   End
   Begin VB.Menu mnuCancel 
      Caption         =   "C&ancel"
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add to Database"
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFull 
         Caption         =   "Search &Full Database"
      End
      Begin VB.Menu mnuSearchSpecific 
         Caption         =   "Search Specific &Record"
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curAmount As Currency
Public transaction As String
Dim curCurrentBalance As Currency
Dim curBalance As Currency
Dim autoNumber As Long

Private Sub cmdAdd_Click()
If txtDate = "" Or txtAmount = "" Or txtDescription = "" Or txtBalance = "" Then
MsgBox "All fields must be entered", vbCritical, "Fill All Fields"
Exit Sub
End If
dbr1.AddNew
dbr1(1) = transaction
dbr1(2) = txtDate.Text
If optTransaction(0) = True Then
dbr1(3) = txtCheckNumber.Text
End If
dbr1(4) = txtDescription.Text
dbr1(5) = curAmount
dbr1(6) = curBalance
dbr1.Update
MsgBox "Record Added", vbInformation
Call Form_Load
End Sub

Private Sub cmdCancel_Click()
Call Form_Load
End Sub

Private Sub cmdCompute_Click()
curAmount = txtAmount.Text
Call OptionTransaction
Call math
If curBalance < 0 Then
   txtBalance.ForeColor = vbRed
   Else
   txtBalance.ForeColor = vbBlack
End If
txtBalance.Text = curBalance
End Sub

Private Sub cmdSearch_Click()
frmSearch.Show
End Sub

Private Sub Form_Load()
Call Connect
txtDate.Text = Date
txtCheckNumber.Text = ""
txtAmount.Text = ""
txtDescription.Text = ""
txtBalance.Text = ""
txtCheckNumber.Visible = False
   lblCheckNumber.Visible = False
   lblAmount.Top = 1800
   txtAmount.Top = 1800
   lblDescription.Top = 2800
   txtDescription.Top = 2800
   lblBalance.Top = 3800
   txtBalance.Top = 3800
   Form1.Height = 6510
txtDate.Text = Date
If dbr1.EOF <> True Then
dbr1.MoveLast
curCurrentBalance = dbr1(6)
Else
curCurrentBalance = 0
End If
txtCurrentBalance.Text = curCurrentBalance
End Sub

Private Sub mnuAdd_Click()
Call cmdAdd_Click
End Sub

Private Sub mnuCancel_Click()
Call cmdCancel_Click
End Sub

Private Sub mnuCompute_Click()
Call cmdCompute_Click
End Sub

Private Sub mnuQuit_Click()
End
End Sub



Private Sub mnuSearchFull_Click()
Call cmdSearch_Click
End Sub

Private Sub mnuSearchSpecific_Click()
If dbr1.EOF <> True Then
Form2.Show
Else
MsgBox "The database is empty", vbOKOnly, "Nothing to Show"
End If
End Sub

Private Sub optTransaction_Click(Index As Integer)
If optTransaction(0) = True Then
   txtCheckNumber.Locked = False
   txtCheckNumber.Visible = True
   lblCheckNumber.Visible = True
   lblAmount.Top = 2800
   txtAmount.Top = 2800
   lblDescription.Top = 3800
   txtDescription.Top = 3800
   lblBalance.Top = 5000
   txtBalance.Top = 5000
   Form1.Height = 6510
   Else
   txtCheckNumber.Visible = False
   lblCheckNumber.Visible = False
   lblAmount.Top = 1800
   txtAmount.Top = 1800
   lblDescription.Top = 2800
   txtDescription.Top = 2800
   lblBalance.Top = 3800
   txtBalance.Top = 3800
   Form1.Height = 6510
End If
End Sub

Public Sub OptionTransaction()
If optTransaction(0) = True Then
    transaction = "Check"
    ElseIf optTransaction(1) = True Then
    transaction = "Check Card"
    ElseIf optTransaction(2) = True Then
    transaction = "Deposit"
   ElseIf optTransaction(3) = True Then
    transaction = "Withdrawal"
End If
End Sub

Public Sub math()
If transaction = "Check" Then
    curBalance = curCurrentBalance - curAmount
    ElseIf transaction = "Check Card" Then
    curBalance = curCurrentBalance - curAmount
    ElseIf transaction = "Deposit" Then
    curBalance = curCurrentBalance + curAmount
    ElseIf transaction = "Withdrawal" Then
    curBalance = curCurrentBalance - curAmount
End If

End Sub




