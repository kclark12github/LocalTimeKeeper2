VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMain 
   Caption         =   "TimeKeeper2"
   ClientHeight    =   4752
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4752
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1140
      TabIndex        =   28
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   972
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   192
      Left            =   0
      TabIndex        =   22
      Top             =   4560
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   339
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10774
            Key             =   "Message"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "1:52 AM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgdGrid 
      Height          =   1812
      Left            =   120
      TabIndex        =   20
      Top             =   2700
      Width           =   7632
      _ExtentX        =   13462
      _ExtentY        =   3196
      _Version        =   393216
      BackColor       =   -2147483626
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Time Sheet for Today"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraEntry 
      Caption         =   "Enter Time Sheet Information..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   2280
      TabIndex        =   13
      Top             =   600
      Width           =   5412
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   4200
         TabIndex        =   29
         Top             =   1698
         Width           =   1092
      End
      Begin MSDataListLib.DataCombo dbcDepartment 
         DataField       =   "DepartmentID"
         Height          =   288
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   3732
         _ExtentX        =   6583
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3060
         TabIndex        =   21
         Top             =   1698
         Width           =   1092
      End
      Begin VB.TextBox txtHours 
         Alignment       =   1  'Right Justify
         DataField       =   "HoursRec"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1560
         TabIndex        =   19
         Text            =   "txtHours"
         Top             =   1680
         Width           =   852
      End
      Begin MSDataListLib.DataCombo dbcProduct 
         DataField       =   "ProductID"
         Height          =   288
         Left            =   1560
         TabIndex        =   24
         Top             =   600
         Width           =   3732
         _ExtentX        =   6583
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbcProject 
         DataField       =   "ProjectID"
         Height          =   288
         Left            =   1560
         TabIndex        =   25
         Top             =   960
         Width           =   3732
         _ExtentX        =   6583
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbcFunction 
         DataField       =   "FunctionID"
         Height          =   288
         Left            =   1560
         TabIndex        =   26
         Top             =   1320
         Width           =   3732
         _ExtentX        =   6583
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblHours 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hours:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   888
         TabIndex        =   18
         Top             =   1728
         Width           =   552
      End
      Begin VB.Label lblFunction 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Function:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   684
         TabIndex        =   17
         Top             =   1368
         Width           =   756
      End
      Begin VB.Label lblProject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Project:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   792
         TabIndex        =   16
         Top             =   1008
         Width           =   648
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Product:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   744
         TabIndex        =   15
         Top             =   648
         Width           =   696
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   420
         TabIndex        =   14
         Top             =   288
         Width           =   1020
      End
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   492
      Left            =   3060
      TabIndex        =   12
      Top             =   60
      Width           =   552
      _ExtentX        =   974
      _ExtentY        =   868
      _Version        =   393216
      FullWidth       =   46
      FullHeight      =   41
   End
   Begin VB.Frame fraHours 
      Caption         =   "Hours Worked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1752
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1992
      Begin VB.TextBox txtTotalTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   660
         TabIndex        =   6
         Text            =   "txtTotalTime"
         Top             =   1380
         Width           =   1212
      End
      Begin VB.TextBox txtLessTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   660
         TabIndex        =   5
         Text            =   "txtLessTime"
         Top             =   1020
         Width           =   1212
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   288
         Left            =   660
         TabIndex        =   3
         Top             =   300
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   24444931
         UpDown          =   -1  'True
         CurrentDate     =   36490
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   288
         Left            =   660
         TabIndex        =   4
         Top             =   660
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   24444931
         UpDown          =   -1  'True
         CurrentDate     =   36490
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   72
         TabIndex        =   10
         Top             =   1440
         Width           =   492
      End
      Begin VB.Label lblLessTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Less:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   108
         TabIndex        =   9
         Top             =   1080
         Width           =   456
      End
      Begin VB.Label lblEndTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "End:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   384
      End
      Begin VB.Label lblStartTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Start:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   444
      End
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   288
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   508
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444928
      CurrentDate     =   36490
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Text            =   "txtUserID"
      Top             =   180
      Width           =   1872
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   3960
      TabIndex        =   11
      Top             =   168
      Width           =   456
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim rsGrid As ADODB.Recordset
Dim rsDepartments As ADODB.Recordset
Dim rsProducts As ADODB.Recordset
Dim rsProjects As ADODB.Recordset
Dim rsFunctions As ADODB.Recordset
Dim rsSummary As ADODB.Recordset
Private Sub cmdCancel_Click()
    On Error Resume Next
    rsGrid.CancelUpdate
    EnableFields False
End Sub
Private Sub cmdEdit_Click()
    If rsGrid.RecordCount > 0 Then
        EnableFields True
    End If
End Sub
Private Sub cmdNew_Click()
    'Will do an rsGrid.AddNew...
    rsGrid.AddNew
    cmdEdit_Click
End Sub
Private Sub cmdUpdate_Click()
    'Will do an rsGrid.Update...
    rsGrid.Update
    cmdCancel_Click
End Sub
Private Sub dbcProduct_Change()
    ResetProjectList
End Sub
Private Sub dtpDate_Change()
    ResetGrid
End Sub
Private Sub EnableFields(fEnabled As Boolean)
    If fEnabled Then
        dbcDepartment.Enabled = True
        dbcDepartment.BackColor = vbWindowBackground
        dbcProduct.Enabled = True
        dbcProduct.BackColor = vbWindowBackground
        dbcProject.Enabled = True
        dbcProject.BackColor = vbWindowBackground
        dbcFunction.Enabled = True
        dbcFunction.BackColor = vbWindowBackground
        txtHours.Enabled = True
        txtHours.BackColor = vbWindowBackground
        cmdUpdate.Visible = True
        cmdCancel.Visible = True
        
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
    Else
        dbcDepartment.Enabled = False
        dbcDepartment.BackColor = vb3DLight
        dbcProduct.Enabled = False
        dbcProduct.BackColor = vb3DLight
        dbcProject.Enabled = False
        dbcProject.BackColor = vb3DLight
        dbcFunction.Enabled = False
        dbcFunction.BackColor = vb3DLight
        txtHours.Enabled = False
        txtHours.BackColor = vb3DLight
        cmdUpdate.Visible = False
        cmdCancel.Visible = False
        
        cmdNew.Enabled = True
        If rsGrid.RecordCount > 0 Then cmdEdit.Enabled = True Else cmdEdit.Enabled = False
    End If
End Sub
Private Sub dtpEndTime_Change()
    txtTotalTime.Text = TimeDiff(dtpStartTime.Value, dtpEndTime.Value)
End Sub
Private Sub dtpStartTime_Change()
    txtTotalTime.Text = TimeDiff(dtpStartTime.Value, dtpEndTime.Value)
End Sub
Private Sub Form_Load()
    txtUserID.Text = "KCLARK"
    Set adoConn = New ADODB.Connection
    adoConn.Open "FileDSN=TimeKeeper2Local.dsn", "Admin", ""
    
    Set rsDepartments = New ADODB.Recordset
    rsDepartments.Open "Select * from [Departments] order by DepartmentID Asc", adoConn, adOpenKeyset, adLockReadOnly
    dbcDepartment.BoundColumn = "DepartmentID"
    dbcDepartment.ListField = "DepartmentID"
    Set dbcDepartment.RowSource = rsDepartments
    
    Set rsProducts = New ADODB.Recordset
    rsProducts.Open "Select distinct ProductID from [Products & Projects] Products order by ProductID Asc", adoConn, adOpenKeyset, adLockReadOnly
    dbcProduct.BoundColumn = "ProductID"
    dbcProduct.ListField = "ProductID"
    Set dbcProduct.RowSource = rsProducts
    
    Set rsProjects = New ADODB.Recordset
    rsProjects.Open "Select distinct ProductID, ProjectID from [Products & Projects] Projects order by ProjectID Asc", adoConn, adOpenKeyset, adLockReadOnly
    dbcProject.BoundColumn = "ProjectID"
    dbcProject.ListField = "ProjectID"
    Set dbcProject.RowSource = rsProjects
    
    Set rsFunctions = New ADODB.Recordset
    rsFunctions.Open "Select * from [Functions] order by FunctionID Asc", adoConn, adOpenKeyset, adLockReadOnly
    dbcFunction.BoundColumn = "FunctionID"
    dbcFunction.ListField = "FunctionID"
    Set dbcFunction.RowSource = rsFunctions
    
    dtpDate.Value = Now
    dtpDate_Change
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CloseRecordset rsGrid, True
    CloseRecordset rsDepartments, True
    CloseRecordset rsProducts, True
    CloseRecordset rsProjects, True
    CloseRecordset rsFunctions, True
    CloseConnection adoConn, True
End Sub
Public Sub ResetGrid()
    Me.MousePointer = vbHourglass
    dgdGrid.Caption = "Time Sheet Entries for " & Format(dtpDate.Value, "Long Date")
    Set dbcDepartment.DataSource = Nothing
    Set dbcProduct.DataSource = Nothing
    Set dbcProject.DataSource = Nothing
    Set dbcFunction.DataSource = Nothing
    Set txtHours.DataSource = Nothing
    Set dgdGrid.DataSource = Nothing
    CloseRecordset rsGrid, True
    
    Set rsGrid = New ADODB.Recordset
    'rsGrid.Open "select EHD.* from [Employee Hours/Day] EHD where EHD.EmployeeID = 'KCLARK' order by EHD.DateRec Desc"
    rsGrid.Open "select EHD.* from [Employee Hours/Day] EHD where EHD.EmployeeID = '" & txtUserID & "' and EHD.DateRec=#" & Format(dtpDate.Value, "Short Date") & "#", adoConn, adOpenKeyset, adLockOptimistic
    Set dgdGrid.DataSource = rsGrid
    dgdGrid.Columns("EmployeeID").Visible = False
    dgdGrid.Columns("DateRec").Visible = False
    dgdGrid.Columns("DepartmentID").Caption = "Department"
    dgdGrid.Columns("ProductID").Caption = "Product"
    dgdGrid.Columns("ProjectID").Caption = "Project"
    dgdGrid.Columns("FunctionID").Caption = "Function"
    dgdGrid.Columns("HoursRec").Caption = "Hours"
    dgdGrid.Columns("Hours").DataFormat.Format = "0.00"
    
    Set dbcDepartment.DataSource = rsGrid
    Set dbcProduct.DataSource = rsGrid
    ResetProjectList
    Set dbcFunction.DataSource = rsGrid
    txtHours.DataField = "HoursRec"
    Set txtHours.DataSource = rsGrid
    
    On Error Resume Next
    Set dtpStartTime.DataSource = Nothing
    Set dtpEndTime.DataSource = Nothing
    Set txtLessTime.DataSource = Nothing
    Set txtTotalTime.DataSource = Nothing
    CloseRecordset rsSummary, True
    Set rsSummary = New ADODB.Recordset
    rsSummary.Open "select * from [Employee Summary] ES where ES.EmployeeID = '" & txtUserID & "' and ES.DateRec=#" & Format(dtpDate.Value, "Short Date") & "#", adoConn, adOpenKeyset, adLockOptimistic
    dtpStartTime.DataField = "Start Time"
    Set dtpStartTime.DataSource = rsSummary
    dtpEndTime.DataField = "End Time"
    Set dtpEndTime.DataSource = rsSummary
    txtLessTime.DataField = "Less Time"
    Set txtLessTime.DataSource = rsSummary
    txtTotalTime.DataField = "Total Time"
    Set txtTotalTime.DataSource = rsSummary
        
    If rsGrid.EOF Then
        dtpStartTime.Value = "09:00 AM"
        dtpEndTime.Value = "05:00 PM"
        txtLessTime.Text = "0"
        txtTotalTime.Text = TimeDiff(dtpStartTime.Value, dtpEndTime.Value)
    End If
    EnableFields False
    Me.MousePointer = vbDefault
End Sub
Public Sub ResetProjectList()
    Set dbcProject.RowSource = Nothing
    CloseRecordset rsProjects, False
    
    rsProjects.Open "Select distinct ProductID, ProjectID from [Products & Projects] Projects where ProductID='" & dbcProduct.BoundText & "' order by ProjectID Asc", adoConn, adOpenKeyset, adLockReadOnly
    Set dbcProject.RowSource = rsProjects
    Set dbcProject.DataSource = rsGrid
End Sub
Private Function TimeDiff(StartDate As Date, EndDate As Date) As String
    Dim Minutes As Long
    Dim Hours As Long
    
    Minutes = DateDiff("n", StartDate, EndDate)
    If Minutes < 0 Then Minutes = Minutes + (24 * 60) 'Assume it wrapped past midnight into the next day...
    Hours = CLng(Minutes \ 60)
    Minutes = CLng(Minutes Mod 60)
    TimeDiff = Format(Hours, "00") & ":" & Format(Minutes, "00")
End Function


