VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMain 
   Caption         =   "TimeKeeper2"
   ClientHeight    =   4752
   ClientLeft      =   48
   ClientTop       =   504
   ClientWidth     =   7800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4752
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timHand 
      Interval        =   1000
      Left            =   2460
      Top             =   120
   End
   Begin VB.Timer timStopWatch 
      Interval        =   1000
      Left            =   3300
      Top             =   120
   End
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
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Edit selected time sheet entry..."
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
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "New time sheet entry..."
      Top             =   2400
      Width           =   972
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   192
      Left            =   0
      TabIndex        =   29
      Top             =   4560
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   339
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
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
            TextSave        =   "3:33 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgdGrid 
      Height          =   1812
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
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
      Caption         =   "Enter TimeKeeper Information..."
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
      Height          =   2052
      Left            =   2340
      TabIndex        =   23
      Top             =   600
      Width           =   5412
      Begin VB.CommandButton cmdDelete 
         Cancel          =   -1  'True
         Caption         =   "&Delete"
         CausesValidation=   0   'False
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
         Left            =   3480
         TabIndex        =   14
         ToolTipText     =   "Delete this entry..."
         Top             =   1698
         Width           =   852
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         CausesValidation=   0   'False
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
         Left            =   4440
         TabIndex        =   15
         ToolTipText     =   "Cancel edit session..."
         Top             =   1698
         Width           =   852
      End
      Begin MSDataListLib.DataCombo dbcDepartment 
         DataField       =   "DepartmentID"
         Height          =   288
         Left            =   1260
         TabIndex        =   8
         ToolTipText     =   "Department against which this time should be ""charged""..."
         Top             =   240
         Width           =   4032
         _ExtentX        =   7112
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
         Left            =   2580
         TabIndex        =   13
         ToolTipText     =   "Update database with current entry..."
         Top             =   1698
         Width           =   852
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
         Left            =   1260
         TabIndex        =   12
         Text            =   "txtHours"
         ToolTipText     =   "Hours spent on this activity..."
         Top             =   1680
         Width           =   852
      End
      Begin MSDataListLib.DataCombo dbcProduct 
         DataField       =   "ProductID"
         Height          =   288
         Left            =   1260
         TabIndex        =   9
         ToolTipText     =   "Product against which this time should be ""charged""..."
         Top             =   600
         Width           =   4032
         _ExtentX        =   7112
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
         Left            =   1260
         TabIndex        =   10
         ToolTipText     =   "Project against which this time should be ""charged""..."
         Top             =   960
         Width           =   4032
         _ExtentX        =   7112
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
         Left            =   1260
         TabIndex        =   11
         ToolTipText     =   "Categorize time spent by these functions..."
         Top             =   1320
         Width           =   4032
         _ExtentX        =   7112
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
         Left            =   588
         TabIndex        =   28
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
         Left            =   384
         TabIndex        =   27
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
         Left            =   492
         TabIndex        =   26
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
         Left            =   444
         TabIndex        =   25
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
         Left            =   120
         TabIndex        =   24
         Top             =   288
         Width           =   1020
      End
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
      TabIndex        =   17
      Top             =   600
      Width           =   1992
      Begin VB.TextBox txtTotalTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         TabIndex        =   5
         Text            =   "txtTotalTime"
         ToolTipText     =   "Total time spent ""on-the-clock""..."
         Top             =   1380
         Width           =   1212
      End
      Begin VB.TextBox txtLessTime 
         Alignment       =   2  'Center
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
         Left            =   660
         TabIndex        =   4
         Text            =   "txtLessTime"
         ToolTipText     =   "Time taken for lunch, breaks, etc."
         Top             =   1020
         Width           =   1212
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   288
         Left            =   660
         TabIndex        =   2
         ToolTipText     =   "Time you ""Clocked-In""..."
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
         TabIndex        =   3
         ToolTipText     =   "Time you ""Clocked-Out""..."
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      TabStop         =   0   'False
      Text            =   "txtUserID"
      ToolTipText     =   "UserID..."
      Top             =   180
      Width           =   1872
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   9
      Left            =   2880
      Picture         =   "frmMain.frx":0442
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   8
      Left            =   2880
      Picture         =   "frmMain.frx":0808
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   7
      Left            =   2880
      Picture         =   "frmMain.frx":0BF1
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   6
      Left            =   2880
      Picture         =   "frmMain.frx":0FD1
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   5
      Left            =   2880
      Picture         =   "frmMain.frx":13A6
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   4
      Left            =   2880
      Picture         =   "frmMain.frx":178B
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   3
      Left            =   2880
      Picture         =   "frmMain.frx":1B6E
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   2
      Left            =   2880
      Picture         =   "frmMain.frx":1F4D
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   1
      Left            =   2880
      Picture         =   "frmMain.frx":2327
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   7
      Left            =   2880
      Picture         =   "frmMain.frx":26F5
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   6
      Left            =   2868
      Picture         =   "frmMain.frx":2B08
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   5
      Left            =   2868
      Picture         =   "frmMain.frx":2F16
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   4
      Left            =   2868
      Picture         =   "frmMain.frx":3328
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   3
      Left            =   2868
      Picture         =   "frmMain.frx":373C
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   2
      Left            =   2868
      Picture         =   "frmMain.frx":3B4F
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   1
      Left            =   2868
      Picture         =   "frmMain.frx":3F5D
      Top             =   60
      Width           =   384
   End
   Begin VB.Image imgHand 
      Height          =   384
      Index           =   0
      Left            =   2880
      Picture         =   "frmMain.frx":4371
      Top             =   60
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgStopWatch 
      Height          =   384
      Index           =   0
      Left            =   2868
      Picture         =   "frmMain.frx":4735
      Top             =   60
      Width           =   384
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "lblA"
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
      Left            =   4020
      TabIndex        =   30
      Top             =   420
      Visible         =   0   'False
      Width           =   336
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
      TabIndex        =   22
      Top             =   168
      Width           =   456
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConn As ADODB.Connection
Dim dtCurrent As Date
Dim fPendingUpdate As Boolean
Dim minDataComboWidth As Single
Dim minDateWidth As Single
Dim minFrameWidth As Single
Dim minGridHeight As Single
Dim minGridWidth As Single
Dim minHeight As Single
Dim minWidth As Single
Dim MouseX As Single
Dim MouseY As Single
Dim rsDepartments As ADODB.Recordset
Dim rsFunctions As ADODB.Recordset
Dim rsProducts As ADODB.Recordset
Dim rsProjects As ADODB.Recordset
Dim rsSummary As ADODB.Recordset
Dim SortDESC() As Boolean
Dim vrsGrid As ADODB.Recordset
Private Sub cmdCancel_Click()
    On Error Resume Next
    
    If vrsGrid.EditMode = adEditInProgress Then vrsGrid.CancelUpdate
    If vrsGrid.EditMode = adEditAdd Then vrsGrid.CancelUpdate
    ResetGrid
    EnableFields False
    dtpDate.Enabled = True
    StopHand
    StartStopWatch
End Sub
Private Sub cmdDelete_Click()
    If vrsGrid.EditMode = adEditInProgress Then vrsGrid.CancelUpdate
    vrsGrid.Delete
    ResetGrid
    EnableFields False
    dtpDate.Enabled = True
    StopHand
    StartStopWatch
End Sub
Private Sub cmdEdit_Click()
    EnableFields True
    cmdDelete.Enabled = True
    dtpDate.Enabled = False
    StopStopWatch
    StartHand
End Sub
Private Sub cmdNew_Click()
    'Will do an vrsGrid.AddNew...
    vrsGrid.AddNew
    EnableFields True
    cmdDelete.Enabled = False
    dtpDate.Enabled = False
    StopStopWatch
    StartHand
End Sub
Private Sub cmdUpdate_Click()
    'Will do an vrsGrid.Update...
    vrsGrid("EmployeeID") = txtUserID.Text
    vrsGrid("DateRec") = Format(dtCurrent, "Short Date")
    vrsGrid("DepartmentID") = dbcDepartment.BoundText
    vrsGrid("ProductID") = dbcProduct.BoundText
    vrsGrid("ProjectID") = dbcProject.BoundText
    vrsGrid("FunctionID") = dbcFunction.BoundText
    vrsGrid.Update
    EnableFields False
    ResizeGrid
    dtpDate.Enabled = True
    StopHand
    StartStopWatch
    fPendingUpdate = True
End Sub
Private Sub dbcDepartment_GotFocus()
    TextSelected
End Sub
Private Sub dbcDepartment_Validate(Cancel As Boolean)
    Dim adoRS As ADODB.Recordset
    Dim RecordsAffected As Long
    
    If Not IsNull(dbcDepartment.SelectedItem) Then Exit Sub
    Set adoRS = New ADODB.Recordset
    adoRS.Open "select DepartmentID from Departments where DepartmentID like '" & dbcDepartment.Text & "%'", adoConn, adOpenKeyset, adLockReadOnly
    If Not adoRS.EOF Then
        dbcDepartment.BoundText = adoRS("DepartmentID")
    Else
        If MsgBox(dbcDepartment.Text & " does not yet exist in the database. Do you want to add this new department?", vbYesNo) = vbYes Then
            adoConn.BeginTrans
            adoConn.Execute "insert into Departments (DepartmentID) values ('" & dbcDepartment.Text & "')", RecordsAffected
            adoConn.CommitTrans
            rsDepartments.Requery
        End If
    End If
    CloseRecordset adoRS, True
End Sub
Private Sub dbcFunction_GotFocus()
    TextSelected
End Sub
Private Sub dbcFunction_Validate(Cancel As Boolean)
    Dim adoRS As ADODB.Recordset
    Dim RecordsAffected As Long
    
    If Not IsNull(dbcFunction.SelectedItem) Then Exit Sub
    Set adoRS = New ADODB.Recordset
    adoRS.Open "select FunctionID from Functions where FunctionID like '" & dbcFunction.Text & "%'", adoConn, adOpenKeyset, adLockReadOnly
    If Not adoRS.EOF Then
        dbcFunction.BoundText = adoRS("FunctionID")
    Else
        If MsgBox(dbcFunction.Text & " does not yet exist in the database. Do you want to add this new Function?", vbYesNo) = vbYes Then
            adoConn.BeginTrans
            adoConn.Execute "insert into Functions (FunctionID) values ('" & dbcFunction.Text & "')", RecordsAffected
            adoConn.CommitTrans
            rsFunctions.Requery
        End If
    End If
    CloseRecordset adoRS, True
End Sub
Private Sub dbcProduct_Change()
    ResetProjectList
End Sub
Private Sub dbcProduct_GotFocus()
    TextSelected
End Sub
Private Sub dbcProduct_Validate(Cancel As Boolean)
    Dim adoRS As ADODB.Recordset
    Dim RecordsAffected As Long
    
    If Not IsNull(dbcProduct.SelectedItem) Then Exit Sub
    Set adoRS = New ADODB.Recordset
    adoRS.Open "select ProductID from [Products & Projects] where ProductID like '" & dbcProduct.Text & "%'", adoConn, adOpenKeyset, adLockReadOnly
    If Not adoRS.EOF Then
        dbcProduct.BoundText = adoRS("ProductID")
    Else
        If MsgBox(dbcProduct.Text & " does not yet exist in the database. Do you want to add this new Product?", vbYesNo) = vbYes Then
            adoConn.BeginTrans
            adoConn.Execute "insert into [Products & Projects] ([ProductID],[ProjectID],[ClientID],[CustomID],[YearRec],[R & D],[InActive],[Regulatory],[SGCustID],[SGFuncID],[SGOptID]) values " & _
                "('" & dbcProduct.Text & "', " & _
                "'General Support', " & _
                "Null, " & _
                "0, " & _
                "Null, " & _
                "0, " & _
                "0, " & _
                "0, " & _
                "Null, " & _
                "Null, " & _
                "Null)", RecordsAffected
            adoConn.CommitTrans
            rsProducts.Requery
        End If
    End If
    CloseRecordset adoRS, True
End Sub
Private Sub dbcProject_GotFocus()
    TextSelected
End Sub
Private Sub dbcProject_Validate(Cancel As Boolean)
    Dim adoRS As ADODB.Recordset
    Dim RecordsAffected As Long
    
    If Not IsNull(dbcProject.SelectedItem) Then Exit Sub
    Set adoRS = New ADODB.Recordset
    adoRS.Open "select ProjectID from [Products & Projects] where ProductID='" & dbcProduct.BoundText & "' And ProjectID like '" & dbcProject.Text & "%'", adoConn, adOpenKeyset, adLockReadOnly
    If Not adoRS.EOF Then
        dbcProject.BoundText = adoRS("ProjectID")
    Else
        If MsgBox(dbcProject.Text & " does not yet exist in the database. Do you want to add this new Project?", vbYesNo) = vbYes Then
            adoConn.BeginTrans
            adoConn.Execute "insert into [Products & Projects] ([ProductID],[ProjectID],[ClientID],[CustomID],[YearRec],[R & D],[InActive],[Regulatory],[SGCustID],[SGFuncID],[SGOptID]) values " & _
                "('" & dbcProduct.Text & "', " & _
                "'" & dbcProject.Text & "', " & _
                "Null, " & _
                "0, " & _
                "Null, " & _
                "0, " & _
                "0, " & _
                "0, " & _
                "Null, " & _
                "Null, " & _
                "Null)", RecordsAffected
            adoConn.CommitTrans
            rsProducts.Requery
        End If
    End If
    CloseRecordset adoRS, True
End Sub
Private Sub dgdGrid_DblClick()
    Dim col As Column
    Dim ColRight As Single
    Dim iCol As Integer
    Dim ResizeWindow As Single
    Dim rsTemp As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    ResizeWindow = 36
    For iCol = dgdGrid.LeftCol To dgdGrid.Columns.Count - 1
        Set col = dgdGrid.Columns(iCol)
        If col.Visible And col.Width > 0 Then
            ColRight = col.Left + col.Width
            If MouseY <= (col.Top + (dgdGrid.RowHeight * dgdGrid.HeadLines)) _
                And MouseX >= (ColRight - ResizeWindow) _
                And MouseX <= (ColRight + ResizeWindow) Then
                dgdGrid.ClearSelCols
                Set rsTemp = vrsGrid.Clone(adLockReadOnly)
                ResizeColumn dgdGrid, rsTemp, col
                CloseRecordset rsTemp, True
                GoTo ExitSub
            End If
        End If
    Next iCol
    
    'If we get here, the user isn't trying to resize a column, so select the row...
    dgdGrid.ClearSelCols
    'If col.Visible And col.Top > 0 And MouseY > col.Top Then cmdOK_Click
    
ExitSub:
    Me.MousePointer = vbDefault
End Sub
Private Sub dgdGrid_HeadClick(ByVal ColIndex As Integer)
    Dim saveBookmark As Variant
    
    On Error Resume Next
    saveBookmark = dgdGrid.Bookmark
    vrsGrid.Sort = vbNullString
    If SortDESC(ColIndex) Then
        vrsGrid.Sort = vrsGrid(ColIndex).Name & " DESC"
    Else
        vrsGrid.Sort = vrsGrid(ColIndex).Name & " ASC"
    End If
    dgdGrid.ClearSelCols
    
    SortDESC(ColIndex) = Not SortDESC(ColIndex)
End Sub
Private Sub dgdGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub
Private Sub dgdGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim i As Long
    
    For i = 0 To dgdGrid.SelBookmarks.Count - 1
        dgdGrid.SelBookmarks.Remove 0
    Next i
    If Not IsNull(dgdGrid.Bookmark) Then dgdGrid.SelBookmarks.Add dgdGrid.Bookmark
    If dgdGrid.Columns.Count > 2 Then dgdGrid.col = dgdGrid.Columns("EmployeeID").ColIndex
End Sub
Private Sub dtpDate_Change()
    'Make sure this uses the not-yet-changed dtCurrent...
    If fPendingUpdate Then UpdateDatabase
    dtCurrent = dtpDate.Value
    ResetGrid
End Sub
Private Sub EnableFields(fEnabled As Boolean)
    If fEnabled Then
        dtpStartTime.Enabled = True
        dtpEndTime.Enabled = True
        txtLessTime.Enabled = True
        txtLessTime.BackColor = vbWindowBackground
        
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
        cmdDelete.Visible = True
        cmdCancel.Visible = True
        
        cmdNew.Enabled = False
        cmdEdit.Enabled = False
    Else
        dtpStartTime.Enabled = False
        dtpEndTime.Enabled = False
        txtLessTime.Enabled = False
        txtLessTime.BackColor = vb3DLight
        
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
        cmdDelete.Visible = False
        cmdCancel.Visible = False
        
        cmdNew.Enabled = True
        If vrsGrid.RecordCount > 0 Then cmdEdit.Enabled = True Else cmdEdit.Enabled = False
    End If
End Sub
Private Sub dtpEndTime_Change()
    txtTotalTime.Text = TimeDiff(dtpStartTime.Value, dtpEndTime.Value)
End Sub
Private Sub dtpStartTime_Change()
    txtTotalTime.Text = TimeDiff(dtpStartTime.Value, dtpEndTime.Value)
End Sub
Private Sub Form_Load()
    minWidth = Me.ScaleWidth
    minHeight = Me.ScaleHeight
    minGridWidth = dgdGrid.Width
    minGridHeight = dgdGrid.Height
    minFrameWidth = fraEntry.Width
    minDataComboWidth = dbcDepartment.Width
    minDateWidth = dtpDate.Width
    
    StartStopWatch
    StopHand
    
    txtUserID.Text = "KCLARK"
    Set adoConn = New ADODB.Connection
    EstablishConnection adoConn, "TimeKeeper2Local.dsn"
    
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
    
    fPendingUpdate = False
    dtCurrent = Now
    dtpDate.Value = dtCurrent
    dtpDate_Change
End Sub
Private Sub Form_Resize()
    Dim xAdjust As Single
    Dim yAdjust As Single
    
    If Me.ScaleWidth < minWidth Or Me.ScaleHeight < minHeight Then Exit Sub
    xAdjust = Me.ScaleWidth - minWidth
    yAdjust = Me.ScaleHeight - minHeight
    
    'Resize the Grid...
    dgdGrid.Height = minGridHeight + yAdjust
    dgdGrid.Width = minGridWidth + xAdjust
    ResizeGrid
    
    'Resize the Data Entry Frame...
    fraEntry.Width = minFrameWidth + xAdjust
    dbcDepartment.Width = minDataComboWidth + xAdjust
    dbcProduct.Width = minDataComboWidth + xAdjust
    dbcProject.Width = minDataComboWidth + xAdjust
    dbcFunction.Width = minDataComboWidth + xAdjust
    cmdCancel.Left = fraEntry.Width - (cmdCancel.Width + 120)
    cmdDelete.Left = cmdCancel.Left - (cmdDelete.Width + 48)
    cmdUpdate.Left = cmdDelete.Left - (cmdUpdate.Width + 48)
    
    dtpDate.Width = minDateWidth + xAdjust
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If fPendingUpdate Then UpdateDatabase
    
    Set dgdGrid.DataSource = Nothing
    Set dbcDepartment.DataSource = Nothing
    Set dbcProduct.DataSource = Nothing
    Set dbcProject.DataSource = Nothing
    Set dbcFunction.DataSource = Nothing
    Set txtHours.DataSource = Nothing
    CloseRecordset vrsGrid, True
    
    Set dbcDepartment.RowSource = Nothing
    CloseRecordset rsDepartments, True
    
    Set dbcProduct.RowSource = Nothing
    CloseRecordset rsProducts, True
    
    Set dbcProject.RowSource = Nothing
    CloseRecordset rsProjects, True
    
    Set dbcFunction.RowSource = Nothing
    CloseRecordset rsFunctions, True
    
    Set dtpStartTime.DataSource = Nothing
    Set dtpEndTime.DataSource = Nothing
    Set txtLessTime.DataSource = Nothing
    Set txtTotalTime.DataSource = Nothing
    CloseRecordset rsSummary, True
    
    CloseConnection adoConn, True
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Public Sub ResetGrid()
    Dim sqlSource As String
    Dim RecordsAffected As Long
    
    Me.MousePointer = vbHourglass
    dgdGrid.Caption = "TimeKeeper Entries for " & Format(dtCurrent, "Long Date")
    Set dbcDepartment.DataSource = Nothing
    Set dbcProduct.DataSource = Nothing
    Set dbcProject.DataSource = Nothing
    Set dbcFunction.DataSource = Nothing
    Set txtHours.DataSource = Nothing
    Set dgdGrid.DataSource = Nothing
    CloseRecordset vrsGrid, True
    
    Set vrsGrid = New ADODB.Recordset
    'vrsGrid.Open "select EHD.* from [Employee Hours/Day] EHD where EHD.EmployeeID = 'KCLARK' order by EHD.DateRec Desc"
    If Not MakeVirtualRecordset(adoConn, "select EHD.* from [Employee Hours/Day] EHD where EHD.EmployeeID = '" & txtUserID & "' and EHD.DateRec=#" & Format(dtCurrent, "Short Date") & "#", vrsGrid) Then
        MsgBox "Problem retrieving data for " & txtUserID.Text & " on " & Format(dtCurrent, "Short Date"), vbCritical, Me.Caption
        End
    End If
    Set dgdGrid.DataSource = vrsGrid
    dgdGrid.Columns("EmployeeID").Visible = False
    dgdGrid.Columns("DateRec").Visible = False
    dgdGrid.Columns("DepartmentID").Caption = "Department"
    dgdGrid.Columns("ProductID").Caption = "Product"
    dgdGrid.Columns("ProjectID").Caption = "Project"
    dgdGrid.Columns("FunctionID").Caption = "Function"
    dgdGrid.Columns("HoursRec").Caption = "Hours"
    dgdGrid.Columns("Hours").DataFormat.Format = "0.00"
    ReDim SortDESC(0 To dgdGrid.Columns.Count - 1)
    
    dgdGrid.AllowRowSizing = False
    dgdGrid.ScrollBars = dbgAutomatic
    dgdGrid.Enabled = True
    dgdGrid.BackColor = vb3DLight
    
    ResizeGrid
    UpdateHoursEntered
    
    Set dbcDepartment.DataSource = vrsGrid
    Set dbcProduct.DataSource = vrsGrid
    ResetProjectList
    Set dbcFunction.DataSource = vrsGrid
    txtHours.DataField = "HoursRec"
    Set txtHours.DataSource = vrsGrid
    
    On Error Resume Next
    Set dtpStartTime.DataSource = Nothing
    Set dtpEndTime.DataSource = Nothing
    Set txtLessTime.DataSource = Nothing
    Set txtTotalTime.DataSource = Nothing
    CloseRecordset rsSummary, True
    Set rsSummary = New ADODB.Recordset
    rsSummary.Open "select * from [Employee Summary] ES where ES.EmployeeID = '" & txtUserID & "' and ES.DateRec=#" & Format(dtCurrent, "Short Date") & "#", adoConn, adOpenKeyset, adLockOptimistic
    If rsSummary.EOF Then
        'Add one...
        adoConn.BeginTrans
        sqlSource = "insert into [Employee Summary] ([EmployeeID],[DateRec],[Start Time],[End Time],[Less Time],[Computed Time],[Total Time],[Comments]) values " & _
            "('" & txtUserID.Text & "', " & "#" & Format(dtCurrent, "Short Date") & "#, " & "'09:00 AM', " & "'05:00 PM', " & "'0.00', " & "'0.00', " & "'0.00', " & "' ')"
        adoConn.Execute sqlSource, RecordsAffected
        adoConn.CommitTrans
        rsSummary.Requery
    End If
    dtpStartTime.DataField = "Start Time"
    Set dtpStartTime.DataSource = rsSummary
    dtpEndTime.DataField = "End Time"
    Set dtpEndTime.DataSource = rsSummary
    txtLessTime.DataField = "Less Time"
    Set txtLessTime.DataSource = rsSummary
    txtTotalTime.DataField = "Total Time"
    Set txtTotalTime.DataSource = rsSummary
        
    If vrsGrid.EOF Then
        dtpStartTime.Value = "09:00 AM"
        dtpEndTime.Value = "05:00 PM"
        txtLessTime.Text = "0.00"
        txtTotalTime.Text = TimeDiff(dtpStartTime.Value, dtpEndTime.Value)
    End If
    EnableFields False
    Me.MousePointer = vbDefault
End Sub
Public Sub ResetProjectList()
    Set dbcProject.DataSource = Nothing
    Set dbcProject.RowSource = Nothing
    CloseRecordset rsProjects, False
    
    rsProjects.Open "Select distinct ProductID, ProjectID from [Products & Projects] Projects where ProductID='" & dbcProduct.BoundText & "' order by ProjectID Asc", adoConn, adOpenKeyset, adLockReadOnly
    If Not rsProjects.EOF Then
        Set dbcProject.RowSource = rsProjects
    Else
        dbcProject.Text = vbNullString
    End If
    Set dbcProject.DataSource = vrsGrid
End Sub
Private Sub ResizeColumn(ctlGrid As Control, rs As ADODB.Recordset, col As Column)
    Dim ColumnFormat As New StdDataFormat
    Dim DataWidth As Long
    Dim ResizeWindow As Single
    Dim WidestData As Long
    
    ResizeWindow = 36
    lblA.Caption = col.Caption
    WidestData = lblA.Width
    Set ColumnFormat = col.DataFormat
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        If Not IsNull(rs(col.ColIndex).Value) Then
            If Not ColumnFormat Is Nothing Then
                lblA.Caption = Format(rs(col.ColIndex).Value, col.DataFormat.Format)
            Else
                lblA.Caption = CStr(rs(col.ColIndex).Value)
            End If
            'Debug.Print "Width of " & lblA.Caption & ": " & lblA.Width
            DataWidth = lblA.Width
            If DataWidth > WidestData Then WidestData = DataWidth
        End If
        rs.MoveNext
    Wend
    Set ColumnFormat = Nothing
    col.Width = WidestData + (4 * ResizeWindow)
    If col.Width > ctlGrid.Width Then col.Width = col.Width - ResizeWindow
End Sub
Private Sub ResizeGrid()
    Dim rsTemp As ADODB.Recordset
    Dim col As Column
    Dim ScaleWidth As Single
    Dim TotalColumnWidths As Single
    
    ScaleWidth = dgdGrid.Width - dgdGrid.Columns("Department").Left - (dgdGrid.Columns.Count * 2)   'This to cover the column delimiter gridlines (I made it up)...
    Set rsTemp = vrsGrid.Clone(adLockReadOnly)
    For Each col In dgdGrid.Columns
        If col.Visible Then
            ResizeColumn dgdGrid, rsTemp, col
            TotalColumnWidths = TotalColumnWidths + col.Width
        End If
    Next col
    CloseRecordset rsTemp, True
    
    If TotalColumnWidths < ScaleWidth Then
        dgdGrid.Columns("Project").Width = dgdGrid.Columns("Project").Width + (ScaleWidth - TotalColumnWidths)
    End If
End Sub
Private Sub StartHand()
    Dim i As Integer
    timHand.Interval = 1000
    timHand.Enabled = True
End Sub
Private Sub StartStopWatch()
    Dim i As Integer
    Dim dInterval As Double
    dInterval = 60000#
    dInterval = dInterval / 8
    timStopWatch.Interval = CInt(dInterval)
    timStopWatch.Enabled = True
    For i = 0 To 7
        imgStopWatch(i).Visible = True
    Next i
End Sub
Private Sub StopHand()
    Dim i As Integer
    timHand.Enabled = False
    For i = 0 To 9
        imgHand(i).Visible = False
    Next i
End Sub
Private Sub StopStopWatch()
    Dim i As Integer
    'timStopWatch.Enabled = False
    For i = 0 To 7
        imgStopWatch(i).Visible = False
    Next i
End Sub
Private Function TimeDiff(StartDate As Date, EndDate As Date) As String
    Const HHMM As Boolean = False
    Dim Minutes As Long
    Dim Hours As Long
    
    If HHMM Then
        Minutes = DateDiff("n", StartDate, EndDate)
        If Minutes < 0 Then Minutes = Minutes + (24 * 60) 'Assume it wrapped past midnight into the next day...
        Hours = CLng(Minutes \ 60)
        Minutes = CLng(Minutes Mod 60)
        TimeDiff = Format(Hours, "00") & ":" & Format(Minutes, "00")
    Else
        Minutes = DateDiff("n", StartDate, EndDate)
        TimeDiff = Format(Minutes / 60, "0.00")
    End If
End Function
Private Sub timHand_Timer()
    Static LastImage As Integer
    Dim i As Integer
    For i = 0 To 9
        imgHand(i).Visible = False
    Next i
    LastImage = LastImage Mod 10
    imgHand(LastImage).ZOrder 0
    imgHand(LastImage).Visible = True
    LastImage = LastImage + 1
End Sub
Private Sub timStopWatch_Timer()
    Static LastImage As Integer
    LastImage = LastImage Mod 8
    imgStopWatch(LastImage).ZOrder 0
    LastImage = LastImage + 1
End Sub
Private Sub txtHours_GotFocus()
    TextSelected
End Sub
Private Sub txtHours_KeyPress(KeyAscii As Integer)
    KeyPressReal KeyAscii
End Sub
Private Sub txtHours_Validate(Cancel As Boolean)
    Dim TotalHours As String
    
    If Trim(txtHours.Text) = vbNullString Then
        MsgBox "Hours must be specified.", vbExclamation, Me.Caption
        Cancel = True
        Exit Sub
    End If
    
    If Val(txtHours.Text) > Val(txtTotalTime.Text) Then
        MsgBox txtHours.Text & " Hours exceeds the specified Total Hours Worked (" & txtTotalTime.Text & ").", vbInformation, Me.Caption
        Cancel = False
        Exit Sub
    End If
    
    'Make sure Grid is updated before calling UpdateHoursEntered...
    vrsGrid("HoursRec") = txtHours.Text
    UpdateHoursEntered
    If sbStatus.Panels("Status").Text <> vbNullString Then
        TotalHours = Mid(sbStatus.Panels("Status").Text, Len("Hours: ") + 1)
        If Val(TotalHours) > Val(txtTotalTime.Text) Then
            MsgBox TotalHours & " Hours exceeds the specified Total Hours Worked (" & txtTotalTime.Text & ").", vbInformation, Me.Caption
            Cancel = False
            Exit Sub
        End If
    End If
End Sub
Private Sub txtLessTime_GotFocus()
    TextSelected
End Sub
Private Sub txtLessTime_KeyPress(KeyAscii As Integer)
    KeyPressReal KeyAscii
End Sub
Private Sub txtLessTime_Validate(Cancel As Boolean)
    If Trim(txtLessTime.Text) = vbNullString Then
        MsgBox "Hours must be specified.", vbExclamation, Me.Caption
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub UpdateDatabase()
    Dim adoRS As ADODB.Recordset
    Dim sqlSource As String
    Dim RecordsAffected As Long
    
    'This routines takes the stuff in the Virtual Recordset, vrsGrid and
    'physically writes it into the database...
    
    On Error GoTo ErrorHandler
    adoConn.BeginTrans
    
    'First replace the Employee Summary record...
    sqlSource = "delete from [Employee Summary] ES where EmployeeID='" & txtUserID.Text & "' And DateRec=#" & Format(dtCurrent, "Short Date") & "#"
    adoConn.Execute sqlSource, RecordsAffected
    sqlSource = "insert into [Employee Summary] ([EmployeeID],[DateRec],[Start Time],[End Time],[Less Time],[Computed Time],[Total Time],[Comments]) values " & _
        "('" & txtUserID.Text & "', " & _
        "#" & Format(dtCurrent, "Short Date") & "#, " & _
        "'" & dtpStartTime.Value & "', " & _
        "'" & dtpEndTime.Value & "', " & _
        "'" & txtLessTime.Text & "', " & _
        "'" & Mid(sbStatus.Panels("Status").Text, Len("Hours: ") + 1) & "', " & _
        "'" & txtTotalTime.Text & "', " & _
        "' ')"
    adoConn.Execute sqlSource, RecordsAffected
    
    'OK, now do the detail...
    sqlSource = "delete from [Employee Hours/Day] EHD where EmployeeID='" & txtUserID.Text & "' And DateRec=#" & Format(dtCurrent, "Short Date") & "#"
    adoConn.Execute sqlSource, RecordsAffected
    If vrsGrid.RecordCount > 0 Then
        vrsGrid.MoveFirst
        While Not vrsGrid.EOF
            sqlSource = "insert into [Employee Hours/Day] ([EmployeeID],[DateRec],[DepartmentID],[ProductID],[ProjectID],[FunctionID],[HoursRec]) values " & _
                "('" & txtUserID.Text & "', " & _
                "#" & Format(dtCurrent, "Short Date") & "#, " & _
                "'" & vrsGrid("DepartmentID") & "', " & _
                "'" & vrsGrid("ProductID") & "', " & _
                "'" & vrsGrid("ProjectID") & "', " & _
                "'" & vrsGrid("FunctionID") & "', " & _
                "'" & vrsGrid("HoursRec") & "')"
            adoConn.Execute sqlSource, RecordsAffected
            vrsGrid.MoveNext
        Wend
    End If
    adoConn.CommitTrans
    
ExitSub:
    fPendingUpdate = False
    Exit Sub
    
ErrorHandler:
    Dim ErrCode As Long
    Dim ErrString As String
    
    ErrString = BuildADOerror(adoConn, ErrCode)
    adoConn.RollbackTrans
    MsgBox ErrString, vbExclamation, Me.Caption
    GoTo ExitSub
    Resume Next 'for debug purposes only...
End Sub
Private Sub UpdateHoursEntered()
    Dim adoRS As ADODB.Recordset
    Dim Hours As Double
    
    Hours = 0#
    If vrsGrid.RecordCount > 0 Then
        Set adoRS = vrsGrid.Clone(adLockReadOnly) 'So user doesn't see us scolling through the recordset in the grid...
        If adoRS.RecordCount > 0 Then
            adoRS.MoveFirst
            While Not adoRS.EOF
                Hours = Hours + adoRS("HoursRec")
                adoRS.MoveNext
            Wend
        End If
        CloseRecordset adoRS, True
    End If
    sbStatus.Panels("Status").Text = "Hours: " & Format(Hours, "0.00")
End Sub

