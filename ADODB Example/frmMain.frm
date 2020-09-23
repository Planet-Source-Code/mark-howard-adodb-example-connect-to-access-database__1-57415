VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADODB Example"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtTitle1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtFirstName1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtSurname1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtPhone1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtPhone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtSurname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFirstName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Title -"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "First Name -"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Surname -"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone -"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4200
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label5 
      Caption         =   "ID -"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone -"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Surname -"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "First Name -"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Title -"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AConnection As New ADODB.Connection
Dim ARecordset As New ADODB.Recordset

Private Sub cmdAdd_Click()
On Error Resume Next
' Check that all fields are filled out -
If txtTitle1.Text = "" Then
MsgBox "Please enter a Title.", vbOKOnly, "Data Required"
Exit Sub
Else:
    If txtFirstName1.Text = "" Then
    MsgBox "Please enter a First Name.", vbOKOnly, "Data Required"
    Exit Sub
    Else:
        If txtSurname1.Text = "" Then
        MsgBox "Please enter a Surname.", vbOKOnly, "Data Required"
        Exit Sub
        Else:
            If txtPhone1.Text = "" Then
            MsgBox "Please enter a Phone Number.", vbOKOnly, "Data Required"
            Exit Sub
            End If
        End If
    End If
End If
' Add new record -
ARecordset.AddNew
ARecordset.Fields("Title") = txtTitle1.Text & " "
ARecordset.Fields("First Name") = txtFirstName1.Text & " "
ARecordset.Fields("Surname") = txtSurname1.Text & " "
ARecordset.Fields("Phone Number") = txtPhone1.Text & " "
' Clear the fields -
txtTitle1.Text = ""
txtFirstName1.Text = ""
txtSurname1.Text = ""
txtPhone1.Text = ""
' Go to the new record -
ARecordset.MoveLast
GetFields
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
' Delete current record -
ARecordset.Delete adAffectCurrent
' Go to the first record -
ARecordset.MoveFirst
GetFields
End Sub

Private Sub cmdFirst_Click()
On Error Resume Next
' Go to first record -
ARecordset.MoveFirst
GetFields
End Sub

Private Sub cmdLast_Click()
On Error Resume Next
' Go to last record -
ARecordset.MoveLast
GetFields
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
' Go to next record -
ARecordset.MoveNext
GetFields
End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
' Go to next record -
ARecordset.MovePrevious
GetFields
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
' Connect to database -
AConnection.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;data source=Database.mdb;"
AConnection.CursorLocation = adUseClient
AConnection.Open
' Open the 'Customers' table -
ARecordset.Open "Select * from Customers", AConnection, adOpenDynamic, adLockOptimistic
If AConnection.State = 1 Then
' If database is connected, fill in the fields -
GetFields
End If
Exit Sub
' Error message -
ErrorHandler:
MsgBox Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "The program will now close.", vbOKOnly, "Error!"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Close the connection when the form is closed -
AConnection.Close
Set AConnection = Nothing
End Sub

Private Sub GetFields()
' Places field data into the text boxes -
txtID.Text = ARecordset(0)
txtTitle.Text = ARecordset(1)
txtFirstName.Text = ARecordset(2)
txtSurname.Text = ARecordset(3)
txtPhone.Text = ARecordset(4)
End Sub

