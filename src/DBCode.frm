VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Code Generator"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "DBCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraDB 
      Enabled         =   0   'False
      Height          =   5445
      Left            =   90
      TabIndex        =   11
      Top             =   960
      Width           =   5025
      Begin VB.CommandButton cmdTag 
         Caption         =   "Untag All"
         Height          =   345
         Index           =   1
         Left            =   3690
         TabIndex        =   6
         Top             =   600
         Width           =   1245
      End
      Begin VB.CommandButton cmdTag 
         Caption         =   "Tag All"
         Height          =   345
         Index           =   0
         Left            =   3690
         TabIndex        =   5
         Top             =   210
         Width           =   1245
      End
      Begin VB.ListBox lstTables 
         Height          =   5130
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   3225
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   345
         Left            =   3690
         TabIndex        =   9
         Top             =   1830
         Width           =   1245
      End
      Begin VB.CommandButton cmdVBCode 
         Caption         =   "VB Code..."
         Height          =   345
         Left            =   3690
         TabIndex        =   7
         Top             =   1050
         Width           =   1245
      End
      Begin VB.CommandButton cmdSQLCode 
         Caption         =   "SQL Script..."
         Height          =   345
         Left            =   3690
         TabIndex        =   8
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   345
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   570
      Width           =   1245
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   210
      Width           =   4185
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Connect"
      Height          =   345
      Left            =   3840
      TabIndex        =   3
      Top             =   570
      Width           =   1245
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "DSN..."
      Height          =   345
      Index           =   0
      Left            =   2550
      TabIndex        =   2
      Top             =   570
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBrowse_Click(Index As Integer)
    Dim db As DAO.Database
    
    On Local Error GoTo BrowseError
    
    If Index = 0 Then
        Set db = DBEngine.OpenDatabase("", False, False, "")
        txtDB.Text = db.Connect
        cmdOpen.Value = True
    Else
        With CommonDialog1
            .CancelError = True
            .DialogTitle = "Select database.."
            .Flags = cdlOFNHideReadOnly
            .Filter = "Access Database (.mdb)|*.mdb"
            .ShowOpen
            
            txtDB.Text = .FileName
            cmdOpen.Value = True
        End With
    End If
    
BrowseError:
    Exit Sub
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    Dim tbl As DAO.TableDef
    Dim sPW As String
    
    On Local Error GoTo OpenError
    
    If Len(txtDB.Text) > 0 Then
        Set mDb = DBEngine.OpenDatabase(txtDB.Text, False, False)
        
        If sPW <> "" Then
            Set mDb = DBEngine.OpenDatabase(txtDB.Text, False, False, "MS Access;PWD=" & sPW)
        End If
        
        Screen.MousePointer = vbHourglass
        
        With lstTables
            .Clear
            
            For Each tbl In mDb.TableDefs
                If (tbl.Attributes And dbSystemObject) Then
                Else
                    .AddItem tbl.Name
                End If
            Next tbl
        End With
        
        fraDB.Enabled = True
        Screen.MousePointer = vbDefault
    Else
        MsgBox "You must specify a data source", vbExclamation
        txtDB.SetFocus
    End If
    Exit Sub

OpenError:
    If Len(sPW) = 0 Then
        sPW = InputBox("Password")
        If Len(sPW) = 0 Then
            Exit Sub
        End If
        Resume Next
    Else
        MsgBox "Invalid Password", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub cmdSQLCode_Click()
    If lstTables.SelCount = 0 Then
        MsgBox "No Tables selected", vbExclamation
        Exit Sub
    End If
    
    frmSQLCode.Show vbModal
End Sub

Private Sub cmdTag_Click(Index As Integer)
    Dim lCount As Long
    
    With lstTables
        If .ListCount > 0 Then
            For lCount = 0 To .ListCount - 1
                .Selected(lCount) = (Index = 0)
            Next lCount
            
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub cmdVBCode_Click()
    If lstTables.SelCount = 0 Then
        MsgBox "No Tables selected", vbExclamation
        Exit Sub
    End If
    
    frmVBCode.Show vbModal
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mDb Is Nothing Then
        mDb.Close
        Set mDb = Nothing
    End If
End Sub


