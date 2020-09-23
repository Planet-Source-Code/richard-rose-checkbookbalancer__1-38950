VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Program"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Done Searching"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Results"
      Height          =   2535
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
      Begin VB.TextBox txtDescription 
         DataField       =   "Description"
         DataSource      =   "Adodc1"
         Height          =   885
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtAmount 
         DataField       =   "Amount"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCheckNumber 
         DataField       =   "CheckNumber"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Date"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   1935
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      Begin VB.OptionButton optSearch 
         Caption         =   "By Description"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "By Amount"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "By Date"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "By Check Number"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlstring As String
Dim tablestring As String
Dim results As Variant
Dim inpCheckNumber As String
Dim inpDate As String
Dim inpAmount As String
Dim inpDescription As String



Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub Form_Load()
For x = 0 To 3
optSearch(x).Value = False
Next x
End Sub

Private Sub optSearch_Click(Index As Integer)
Select Case Index
    Case 0
        inpCheckNumber = InputBox("Enter Check Number", "Check Number")
        If inpCheckNumber = "" Then
        Call Form_Load
        Exit Sub
        End If
        dbr1.MoveFirst
        Do Until dbr1(3) = inpCheckNumber
        dbr1.MoveNext
        If dbr1.EOF = True Then
        MsgBox "Check number " & inpCheckNumber & " was not found", vbOKOnly, "Oops, No Match"
        Call Form_Load
        Exit Sub
        End If
        Loop
        Call FillFields
    Case 1
        inpDate = InputBox("Enter Date", "Date")
        If inpDate = "" Then
        Call Form_Load
        Exit Sub
        End If
        dbr1.MoveFirst
        Do Until dbr1(2) = inpDate
        dbr1.MoveNext
        If dbr1.EOF = True Then
        MsgBox "Date " & inpDate & " was not found", vbOKOnly, "Oops, No Match"
        Call Form_Load
        Exit Sub
        End If
        Loop
        Call FillFields
    Case 2
        inpAmount = InputBox("Enter Amount", "Amount")
        If inpAmount = "" Then
        Call Form_Load
        Exit Sub
        End If
        dbr1.MoveFirst
        Do Until dbr1(5) = inpAmount
        dbr1.MoveNext
        If dbr1.EOF = True Then
        MsgBox "Amount $" & inpAmount & " was not found", vbOKOnly, "Oops, No Match"
        Call Form_Load
        Exit Sub
        End If
        Loop
        Call FillFields
    Case 3
        inpDescription = InputBox("Enter Description", "Description")
        If inpDescription = "" Then
        Call Form_Load
        Exit Sub
        End If
        dbr1.MoveFirst
        Do Until dbr1(4) = inpDescription
        dbr1.MoveNext
        If dbr1.EOF = True Then
        MsgBox "Description " & inpDescription & " was not found", vbOKOnly, "Oops, No Match"
        Call Form_Load
        Exit Sub
        End If
        Loop
        Call FillFields
End Select
End Sub

Public Sub FillFields()
txtDate = dbr1(2)
txtCheckNumber = dbr1(3)
txtAmount = dbr1(5)
txtDescription = dbr1(4)
Call Form_Load
End Sub
