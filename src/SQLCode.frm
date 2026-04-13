VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQLCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Script Generator"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3390
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   1065
      Left            =   90
      TabIndex        =   5
      Top             =   960
      Width           =   2655
      Begin VB.CheckBox chkDropTable 
         Caption         =   "DROP TABLE Statement"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   750
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox chkIndexes 
         Caption         =   "Indexes"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox chkFieldProperties 
         Caption         =   "Default Field Values"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   510
         Value           =   1  'Checked
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Format"
      Height          =   795
      Left            =   90
      TabIndex        =   2
      Top             =   150
      Width           =   2655
      Begin VB.OptionButton optMode 
         Caption         =   "Create single script"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Create Script per Table"
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSQLCode"
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
    Dim idx As DAO.Index
    Dim tbl As DAO.TableDef
    
    Dim lIndex As Long
    Dim lTableCount As Long
    Dim nHandle As Integer
    Dim sPath As String
    Dim sTab As String
    Dim sBuffer As String
    Dim sText As String
    Dim bMemo As Boolean
    
    On Local Error GoTo GenerateError
    
    lIndex = InStrRev(frmMain.txtDB.Text, "\")
    sPath = Left$(frmMain.txtDB.Text, lIndex)
    sTab = Space$(4)
    
    If optMode(1).Value Then
        With CommonDialog1
            .CancelError = True
            .DialogTitle = "Select script.."
            .Flags = cdlOFNHideReadOnly
            .Filter = "SQL Script (.sql)|*.sql|All files (*.*)|*.*"
            .InitDir = sPath
            .ShowSave
            
            nHandle = FreeFile
            Open .FileName For Output As #nHandle
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
                
                If optMode(0).Value Then
                    nHandle = FreeFile
                    Open sPath & tbl.Name & ".sql" For Output As #nHandle
                End If
                
                lTableCount = lTableCount + 1
                
                Print #nHandle, "-- " & String(60, "*")
                
                If chkDropTable.Value Then
                    Print #nHandle, "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & tbl.Name & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
                    Print #nHandle, "drop table [dbo].[" & tbl.Name & "]"
                    Print #nHandle, "GO"
                    Print #nHandle, ""
                End If
                
                Print #nHandle, "CREATE TABLE [dbo].[" & tbl.Name & "] ("
                
                bMemo = False
                sBuffer = ""
                
                For Each fld In tbl.Fields
                    If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
                        sText = "[" & fld.Name & "] [int] IDENTITY (1,1) NOT NULL"
                    Else
                        Select Case fld.Type
                        Case dbText
                            sText = "[" & fld.Name & "] [nvarchar] (" & fld.Size & ") COLLATE Latin1_General_CI_AS"
                        Case dbInteger
                            sText = "[" & fld.Name & "] [smallint]"
                        Case dbLong
                            sText = "[" & fld.Name & "] [int]"
                        Case dbCurrency
                            sText = "[" & fld.Name & "] [money]"
                        Case dbSingle, dbDouble
                            sText = "[" & fld.Name & "] [float]"
                        Case dbDate
                            sText = "[" & fld.Name & "] [smalldatetime]"
                        Case dbBoolean
                            sText = "[" & fld.Name & "] [bit]"
                        
                        Case dbMemo
                            bMemo = True
                            sText = "[" & fld.Name & "] [ntext] COLLATE Latin1_General_CI_AS"
                        End Select
                        
                        If fld.Required Then
                            sText = sText & " NOT NULL"
                        Else
                            sText = sText & " NULL"
                        End If
                    End If
                    
                    If Len(sBuffer) = 0 Then
                        sBuffer = sTab & sText
                    Else
                        sBuffer = sBuffer & " ," & vbCrLf & sTab & sText
                    End If
                Next fld
                
                Print #nHandle, sBuffer
                
                If bMemo Then
                    Print #nHandle, ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"
                Else
                    Print #nHandle, ") ON [PRIMARY]"
                End If
                Print #nHandle, "GO"
                Print #nHandle, ""
                
                If chkIndexes.Value Then
                    For Each idx In tbl.Indexes
                        If idx.Primary Then
                            Print #nHandle, "ALTER TABLE [dbo].[" & tbl.Name & "] WITH NOCHECK ADD"
                            Print #nHandle, "CONSTRAINT [PK_" & tbl.Name & "] PRIMARY KEY  CLUSTERED"
                            Print #nHandle, "("
                            
                            sBuffer = ""
                            For Each fld In idx.Fields
                                If Len(sBuffer) = 0 Then
                                    sBuffer = "[" & fld.Name & "]"
                                Else
                                    sBuffer = sBuffer & " ," & vbCrLf & "[" & fld.Name & "]"
                                End If
                            Next fld
                            
                            Print #nHandle, sBuffer
                            Print #nHandle, ") ON [PRIMARY]"
                            Print #nHandle, "GO"
                            Print #nHandle, ""
                        Else
                            Print #nHandle, "CREATE NONCLUSTERED INDEX [IK_" & idx.Name & "] ON [dbo].[" & tbl.Name & "]"
                            Print #nHandle, "("
                            
                            sBuffer = ""
                            For Each fld In idx.Fields
                                If Len(sBuffer) = 0 Then
                                    sBuffer = "[" & fld.Name & "]"
                                Else
                                    sBuffer = sBuffer & " ," & vbCrLf & "[" & fld.Name & "]"
                                End If
                            Next fld
                            
                            Print #nHandle, sBuffer
                            Print #nHandle, ") ON [PRIMARY]"
                            Print #nHandle, "GO"
                            Print #nHandle, ""
                        End If
                    Next idx
                End If
                
                If chkFieldProperties.Value Then
                    sBuffer = ""
                
                    For Each fld In tbl.Fields
                        If (fld.Attributes And dbAutoIncrField) = dbAutoIncrField Then
                            'No default allowed
                        Else
                            Select Case fld.Type
                            Case dbText
                                sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT '' FOR [" & fld.Name & "]"
                            Case dbInteger, dbLong, dbCurrency, dbSingle, dbDouble
                                If Val(fld.DefaultValue) <> 0 Then
                                    sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (" & fld.DefaultValue & ") FOR [" & fld.Name & "]"
                                Else
                                    sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                                End If
                            Case dbBoolean
                                If Val(fld.DefaultValue) <> 0 Then
                                    sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (1) FOR [" & fld.Name & "]"
                                Else
                                    sText = "CONSTRAINT [DF_" & tbl.Name & "_" & fld.Name & "] DEFAULT (0) FOR [" & fld.Name & "]"
                                End If
                            Case Else
                                sText = ""
                            End Select
                        
                            If Len(sText) > 0 Then
                                If Len(sBuffer) = 0 Then
                                    sBuffer = sTab & sText
                                Else
                                    sBuffer = sBuffer & " ," & vbCrLf & sTab & sText
                                End If
                            End If
                        End If
                    Next fld
                
                    If Len(sBuffer) > 0 Then
                        Print #nHandle, "ALTER TABLE [dbo].[" & tbl.Name & "] WITH NOCHECK ADD"
                        Print #nHandle, sBuffer
                        Print #nHandle, "GO"
                        Print #nHandle, ""
                    End If
                End If
                
                If optMode(0).Value Then
                    Close #nHandle
                End If
            End If
            
            DoEvents
            ProgressBar1.Value = lIndex
        Next lIndex
        
        If optMode(1).Value Then
            Close #nHandle
        End If
        
        Screen.MousePointer = vbDefault
        ProgressBar1.Visible = False
    End With
    Exit Sub

GenerateError:
    Exit Sub
End Sub



