VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "TimeKeeper2"
   ClientHeight    =   4752
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4752
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   192
      Left            =   0
      TabIndex        =   26
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
            TextSave        =   "11:13 PM"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgdTimeSheet 
      Height          =   1812
      Left            =   120
      TabIndex        =   24
      Top             =   2700
      Width           =   7632
      _ExtentX        =   13462
      _ExtentY        =   3196
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
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
      Height          =   2052
      Left            =   2280
      TabIndex        =   13
      Top             =   600
      Width           =   5412
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   252
         Left            =   3600
         TabIndex        =   25
         Top             =   1698
         Width           =   792
      End
      Begin VB.TextBox txtHours 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   1560
         TabIndex        =   23
         Text            =   "txtHours"
         Top             =   1680
         Width           =   852
      End
      Begin VB.ComboBox cboProject 
         Height          =   288
         Left            =   1560
         TabIndex        =   21
         Top             =   960
         Width           =   3732
      End
      Begin VB.ComboBox cboFunction 
         Height          =   288
         Left            =   1572
         TabIndex        =   19
         Top             =   1320
         Width           =   3732
      End
      Begin VB.ComboBox cboProduct 
         Height          =   288
         Left            =   1548
         TabIndex        =   16
         Top             =   600
         Width           =   3732
      End
      Begin VB.ComboBox cboDepartment 
         Height          =   288
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   3732
      End
      Begin VB.Label lblHours 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hours"
         Height          =   192
         Left            =   1008
         TabIndex        =   22
         Top             =   1728
         Width           =   432
      End
      Begin VB.Label lblFunction 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   192
         Left            =   864
         TabIndex        =   20
         Top             =   1368
         Width           =   600
      End
      Begin VB.Label lblProject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Project"
         Height          =   192
         Left            =   948
         TabIndex        =   18
         Top             =   1008
         Width           =   504
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Product"
         Height          =   192
         Left            =   888
         TabIndex        =   17
         Top             =   648
         Width           =   552
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Height          =   192
         Left            =   612
         TabIndex        =   15
         Top             =   288
         Width           =   840
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
      Enabled         =   -1  'True
      FullWidth       =   46
      FullHeight      =   41
   End
   Begin VB.Frame fraHours 
      Caption         =   "Hours Worked"
      Height          =   1752
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1812
      Begin VB.TextBox txtTotalTime 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   660
         TabIndex        =   6
         Text            =   "txtTotalTime"
         Top             =   1380
         Width           =   1032
      End
      Begin VB.TextBox txtLessTime 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   660
         TabIndex        =   5
         Text            =   "txtLessTime"
         Top             =   1020
         Width           =   1032
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   288
         Left            =   660
         TabIndex        =   3
         Top             =   300
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   508
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   24510467
         UpDown          =   -1  'True
         CurrentDate     =   36490
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   288
         Left            =   660
         TabIndex        =   4
         Top             =   660
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   508
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   24510467
         UpDown          =   -1  'True
         CurrentDate     =   36490
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   192
         Left            =   192
         TabIndex        =   10
         Top             =   1440
         Width           =   372
      End
      Begin VB.Label lblLessTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Less"
         Height          =   192
         Left            =   216
         TabIndex        =   9
         Top             =   1080
         Width           =   348
      End
      Begin VB.Label lblEndTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "End"
         Height          =   192
         Left            =   276
         TabIndex        =   8
         Top             =   720
         Width           =   288
      End
      Begin VB.Label lblStartTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Start"
         Height          =   192
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   324
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
      Format          =   24510464
      CurrentDate     =   36490
   End
   Begin VB.TextBox txtUserID 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Text            =   "txtUserID"
      Top             =   180
      Width           =   1872
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   192
      Left            =   3960
      TabIndex        =   11
      Top             =   168
      Width           =   348
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtpDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub
