VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVBCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Code Generator"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   975
      Left            =   90
      TabIndex        =   10
      Top             =   2190
      Width           =   2655
      Begin VB.CheckBox chkIndexes 
         Caption         =   "Indexes"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox chkFieldProperties 
         Caption         =   "Field Properties"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   510
         Value           =   1  'Checked
         Width           =   2385
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Code Format"
      Height          =   975
      Left            =   90
      TabIndex        =   6
      Top             =   150
      Width           =   2655
      Begin VB.OptionButton optMethod 
         Caption         =   "DAO Objects"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "DAO SQL"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2025
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "ADO Objects"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Format"
      Height          =   975
      Left            =   90
      TabIndex        =   2
      Top             =   1170
      Width           =   2655
      Begin VB.OptionButton optMode 
         Caption         =   "Create single Sub"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   690
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Create Sub per Table"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2355
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Create Module per Table"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   660
      Width           =   1275
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   2820
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3450
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVBCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdGenerate_Click()
    Dim fld As DAO.Field
    Dim fldLU As DAO.Field
    Dim idx As DAO.Index
    Dim tbl As DAO.TableDef
    Dim sBuffer As String
    Dim sText As String
    Dim nCodeFormat As Integer
    Dim lIndex As Long
    Dim lTableCount As Long
    Dim bModule As Boolean
    Dim bSub As Boolean
    Dim nHandle As Integer
    Dim sPath As String
    Dim sTab As String
    
    On Local Error GoTo GenerateError
    
    For nCodeFormat = optMethod.LBound To optMethod.UBound
        If optMethod(nCodeFormat).Value Then
            Exit For
        End If
    Next nCodeFormat
    
    bModule = optMode(0).Value
    bSub = optMode(1).Value
        
    lIndex = InStrRev(frmMain.txtDB.Text, "\")
    sPath = Left$(frmMain.txtDB.Text, lIndex)
    sTab = Space$(4)
    
    If Not bModule Then
        With CommonDialog1
            .CancelError = True
            .DialogTitle = "Select module.."
            .Flags = cdlOFNHideReadOnly
            .Filter = "VB Module (.bas)|*.sql|All files (*.*)|*.*"
            .InitDir = sPath
            .ShowOpen
            
            nHandle = FreeFile
            Open .FileName For Output As #nHandle
            
            If Not bSub Then
                Print #nHandle, "Sub CreateTables (DB As DAO.Database)"
            End If
        End With
    End If
    
    With frmMain.lstTables
        Screen.MousePointer = vbHourglass
        
        ProgressBar1.Visible = True
        ProgressBar1.Max = .ListCount
        ProgressBar1.Value = ProgressBar1.Min
        
        For lIndex = 0 To .ListCount - 1
            If .Selected(lIndex) Then
                Set tbl = mDb.TableDefs(.List(lIndex))
                
                If bModule Then
                    nHandle = FreeFile
                    Open sPath & tbl.Name & ".bas" For Output As #nHandle
                    Print #nHandle, "Sub CreateTable(DB As DAO.Database)"
                ElseIf bSub Then
                    Print #nHandle, "Sub " & tbl.Name & "(DB As DAO.Database)"
                End If
                
                If bModule Or bSub Or (lTableCount = 0) Then
                    Select Case nCodeFormat
                    Case 0
                        Print #nHandle, sTab & "Dim fld As DAO.Field"
                        Print #nHandle, sTab & "Dim idx As DAO.Index"
                        Print #nHandle, sTab & "Dim tbl As DAO.TableDef"
                    Case 1
                        Print #nHandle, sTab & "Dim SQL As String"
                    End Select
                    Print #nHandle, ""
                    Print #nHandle, sTab & "'Source Database: " & frmMain.txtDB.Text
                    Print #nHandle, ""
                End If
                
                lTableCount = lTableCount + 1
                
                Print #nHandle, sTab & "'" & String(80, "*")
                Print #nHandle, sTab & "'Code to generate Table: " & tbl.Name
                
                Select Case nCodeFormat
                Case 0
                    Print #nHandle, sTab & "Set tbl = DB.CreateTableDef(" & Chr$(34) & tbl.Name & Chr$(34) & ")"
                    Print #nHandle, sTab & "With tbl"
                    
                    For Each fld In tbl.Fields
                        If (fld.Type = dbText) Or (fld.Type = dbMemo) Then
                            Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & fld.Name & Chr$(34) & "," & GetDBConstant(fld.Type) & "," & fld.Size & ")"
                        Else
                            Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & fld.Name & Chr$(34) & "," & GetDBConstant(fld.Type) & ")"
                        End If
                        
                        If chkFieldProperties.Value Then
                            If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
                                Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & fld.Name & Chr$(34) & ").Attributes = " & fld.Attributes
                            End If
                            If fld.AllowZeroLength Then
                                Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & fld.Name & Chr$(34) & ").AllowZeroLength = True"
                            End If
                            If Len(fld.DefaultValue) > 0 Then
                                If fld.DefaultValue = Chr$(34) & Chr$(34) Then
                                    Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & fld.Name & Chr$(34) & ").DefaultValue = Chr$(34) & Chr$(34)"
                                Else
                                    Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & fld.Name & Chr$(34) & ").DefaultValue = " & fld.DefaultValue
                                End If
                            End If
                            If fld.Required Then
                                Print #nHandle, sTab & sTab & ".Fields(" & Chr$(34) & fld.Name & Chr$(34) & ").Required = True"
                            End If
                        End If
                    Next fld
                    
                    Print #nHandle, sTab & "End With"
                    Print #nHandle, ""
                    Print #nHandle, sTab & "DB.TableDefs.Append tbl"
                    
                    If chkIndexes.Value Then
                        For Each idx In tbl.Indexes
                            Print #nHandle, ""
                            Print #nHandle, sTab & "Set idx = tbl.CreateIndex(" & Chr$(34) & idx.Name & Chr$(34) & ")"
                            Print #nHandle, sTab & "With idx"
                            If idx.Primary Then
                                Print #nHandle, sTab & sTab & ".Primary = True"
                            End If
                            If idx.Unique Then
                                Print #nHandle, sTab & sTab & ".Unique = True"
                            End If
                            
                            For Each fld In idx.Fields
                                Set fldLU = tbl.Fields(fld.Name)
                                
                                If (fldLU.Type = dbText) Or (fldLU.Type = dbMemo) Then
                                    Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & fld.Name & Chr$(34) & "," & GetDBConstant(fldLU.Type) & "," & fldLU.Size & ")"
                                Else
                                    Print #nHandle, sTab & sTab & ".Fields.Append .CreateField(" & Chr$(34) & fld.Name & Chr$(34) & "," & GetDBConstant(fldLU.Type) & ")"
                                End If
                            Next fld
                            
                            Print #nHandle, sTab & "End With"
                            Print #nHandle, sTab & "tbl.Indexes.Append idx"
                        Next idx
                    End If
                
                Case 1
                    Print #nHandle, sTab & "SQL = " & Chr$(34) & "CREATE TABLE " & tbl.Name & " (" & Chr$(34)
                    
                    sBuffer = ""
                    For Each fld In tbl.Fields
                        If (fld.Type = dbText) Then
                            sText = sTab & "SQL = SQL & " & Chr$(34) & fld.Name & " TEXT (" & fld.Size & ")"
                        Else
                            sText = sTab & "SQL = SQL & " & Chr$(34) & fld.Name & " " & GetSQLType(fld.Type)
                        End If
                        
                        If Len(sBuffer) > 0 Then
                            sBuffer = sBuffer & "," & Chr$(34) & vbCrLf & sText
                        Else
                            sBuffer = sText
                        End If
                    Next fld
                    
                    Print #nHandle, sBuffer & ")"
                    Print #nHandle, sTab & "DB.Execute SQL"
                    Print #nHandle, ""
                    
                    If chkIndexes.Value Then
                        For Each idx In tbl.Indexes
                            sBuffer = ""
                            For Each fld In idx.Fields
                                If Len(sBuffer) > 0 Then
                                    sBuffer = sBuffer & "," & fld.Name
                                Else
                                    sBuffer = fld.Name
                                End If
                            Next fld
                            
                            If idx.Primary Then
                                Print #nHandle, sTab & "DB.Execute " & Chr$(34) & "CREATE UNIQUE INDEX " & idx.Name & " ON " & tbl.Name & "(" & sBuffer & ") WITH PRIMARY"
                            ElseIf idx.Unique Then
                                Print #nHandle, sTab & "DB.Execute " & Chr$(34) & "CREATE UNIQUE INDEX " & idx.Name & " ON " & tbl.Name & "(" & sBuffer & ")"
                            Else
                                Print #nHandle, sTab & "DB.Execute " & Chr$(34) & "CREATE INDEX " & idx.Name & " ON " & tbl.Name & "(" & sBuffer & ")"
                            End If
                        Next idx
                    End If
                End Select
                
                Print #nHandle, sTab & "'" & String(80, "*")
                
                If bModule Or bSub Then
                    Print #nHandle, "End Sub"
                Else
                    Print #nHandle, ""
                End If
                
                If bModule Then
                    Close #nHandle
                End If
            End If
            
            DoEvents
            ProgressBar1.Value = lIndex
        Next lIndex
        
        If Not bModule Then
            If Not bSub Then
                Print #nHandle, "End Sub"
            End If
            Close #nHandle
        End If
        
        Screen.MousePointer = vbDefault
        ProgressBar1.Visible = False
    End With
    Exit Sub

GenerateError:
    Exit Sub
End Sub


