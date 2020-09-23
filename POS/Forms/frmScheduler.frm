VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmScheduler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frmScheduler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFilter 
      Height          =   315
      ItemData        =   "frmScheduler.frx":038A
      Left            =   120
      List            =   "frmScheduler.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtSrchStr 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   2040
      TabIndex        =   23
      Text            =   "Search"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Frame freScheduler 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   3
      Left            =   5160
      TabIndex        =   19
      Top             =   240
      Width           =   615
      Begin VB.TextBox txtSchedRmrks 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   645
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtSchedTitle 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpSchedDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54591491
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtSchedDate 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.MonthView mvwSched 
         Height          =   2370
         Left            =   2760
         TabIndex        =   22
         Top             =   0
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483644
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollRate      =   1
         ShowToday       =   0   'False
         StartOfWeek     =   54591489
         TitleBackColor  =   16744448
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483632
         CurrentDate     =   40099
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   29
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Frame freScheduler 
      BorderStyle     =   0  'None
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   5055
      Begin VB.TextBox txtSupplier 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboSchedules 
         Height          =   315
         ItemData        =   "frmScheduler.frx":038E
         Left            =   1800
         List            =   "frmScheduler.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpLastDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   33
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54591491
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtLast 
         Height          =   285
         Left            =   1800
         TabIndex        =   34
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin ctrlButton.ThemedButton cmdViewSup 
         Height          =   375
         Left            =   3360
         TabIndex        =   42
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmScheduler.frx":0392
         Picture         =   "frmScheduler.frx":056C
         PictureAlign    =   2
         PictureSize     =   0
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "* Last Delivery:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Schedules:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblPersonel 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1800
         TabIndex        =   17
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Personel:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame freScheduler 
      BorderStyle     =   0  'None
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   855
      Begin VB.TextBox txtRemarks 
         Height          =   495
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtDescription 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtGapVal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox cboGap 
         Height          =   315
         ItemData        =   "frmScheduler.frx":08C0
         Left            =   1800
         List            =   "frmScheduler.frx":08D3
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* Gap Value:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* Gap:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   525
      End
   End
   Begin ComctlLib.TabStrip tabScheduler 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4895
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delivery Schedules"
            Key             =   "tblDeliverySched"
            Object.Tag             =   "tblSuppliers"
            Object.ToolTipText     =   "sched_type"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Set Supplier Delivery"
            Key             =   "tblSuppliers"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Task Schedules"
            Key             =   "tblSchedules"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lstRecords 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin CtrlLine.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   1800
      TabIndex        =   25
      Top             =   4560
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   4440
      TabIndex        =   36
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":08F7
      Picture         =   "frmScheduler.frx":0AD1
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdView 
      Height          =   375
      Left            =   3120
      TabIndex        =   37
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&View >>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":0E25
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   38
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":0FFF
      Picture         =   "frmScheduler.frx":11D9
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   1800
      TabIndex        =   39
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":152D
      Picture         =   "frmScheduler.frx":1707
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   4440
      TabIndex        =   40
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":1A5B
      Picture         =   "frmScheduler.frx":1C35
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   3120
      TabIndex        =   41
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":1F89
      Picture         =   "frmScheduler.frx":2163
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Indecates required field"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   31
      Top             =   6960
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmScheduler.frx":24B7
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4920
      Picture         =   "frmScheduler.frx":2D81
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SCHEDULER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmScheduler.frx":364B
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save schedules for supplier deliveries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   2310
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sDate As Date

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrMsg
Select Case tabScheduler.SelectedItem.Index
    Case 1
        If MsgBox("Are you sure you want to delete this schedule setup?", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        If lstRecords.SelectedItem.SubItems(1) = "(none)" Then
            MsgBox "You cannot delete this reocord. It is a system record.", vbCritical
            Exit Sub
        End If
        RunSql "Delete * from tblDeliverySched where description = '" & lstRecords.SelectedItem.SubItems(1) & "'"
        MsgBox "Delivery Schedule " & lstRecords.SelectedItem.SubItems(1) & " has been deleted.", vbInformation
    Case 2
        RunSql "Select sched_type from tblSuppliers where record_id = '" & lstRecords.SelectedItem.SubItems(1) & "'"
        Rs.Fields(0) = "(none)"
        Rs.Update
    Case 3
        If MsgBox("Are you sure you want to remove this record?", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        RunSql "delete * from tblSchedules where record_no = " & lstRecords.SelectedItem
        MsgBox "Task schedule has been deleted.", vbInformation
End Select
ClrFlds
Exit Sub
ErrMsg:
    MsgBox ExecErr("Failed to delete record. This record may be used on other related records.", _
            "record_id", tabScheduler.SelectedItem.Tag, tabScheduler.SelectedItem.ToolTipText, lstRecords.SelectedItem.SubItems(1)), vbCritical
End Sub

Private Sub cmdLoad_Click()
lstRecords_DblClick
End Sub

Private Sub cmdNew_Click()
ClrFlds
End Sub

Private Sub cmdview_Click()
If cmdView.Caption = "&View >>" Then
    Me.Height = 7815
    cmdView.Caption = "&View <<"
Else
    Me.Height = 4800
    cmdView.Caption = "&View >>"
End If
End Sub

Private Sub cmdSave_Click()
Dim Record As String

If tabScheduler.SelectedItem.Index = 1 Then
    If TxtEmp(txtDescription, tabScheduler, 1) = True Then Exit Sub
    If CboEmp(cboGap, tabScheduler, 1) = True Then Exit Sub
    If TxtEmp(txtGapVal, tabScheduler, 1) = True Then Exit Sub
ElseIf tabScheduler.SelectedItem.Index = 2 Then
    If txtSupplier.Text = Empty Then
        MsgBox "Please search for a supplier and click Load.", vbExclamation
        Exit Sub
    End If
    If TxtEmp(txtLast) = True Then Exit Sub
Else
    If TxtEmp(txtSchedTitle) = True Then Exit Sub
    If TxtEmp(txtSchedDate) = True Then Exit Sub
End If

If NoRcrd(lstRecords) = True Then
    Record = ""
Else
    Record = lstRecords.SelectedItem.SubItems(1)
End If

RunSql "Select * from " & tabScheduler.SelectedItem.Key & " where " & lstRecords.ColumnHeaders(2).Text & " = '" & Record & "'"
With Rs
    If cmdSave.Caption <> "&Update" Then
        If tabScheduler.SelectedItem.Index = 1 Then
            SubSql "Select * from tblDeliverySched where description = '" & txtDescription.Text & "'"
            If SubRs.EOF = False Then
                MsgBox "Schedule description is already on your database", vbExclamation
                txtDescription.SetFocus
                SelAll txtDescription
                Exit Sub
            End If
            .AddNew
            .Fields!record_no = Val(RcrdId(tabScheduler.SelectedItem.Key, , "record_no"))
            msg = "Added new delivery schedule setting on database."
        ElseIf tabScheduler.SelectedItem.Index = 2 Then
            msg = "Delivery schedule has been updated."
        Else
            .AddNew
            .Fields!record_no = Val(RcrdId(tabScheduler.SelectedItem.Key, , "record_no"))
            msg = "New schedule add on system Scheduler."
        End If
    Else
        msg = "Schedule of " & lstRecords.SelectedItem.SubItems(1) & " has been updated."
    End If
    Select Case tabScheduler.SelectedItem.Index
        Case 1
            .Fields!Description = txtDescription.Text
            .Fields!Gap = cboGap.Text
            .Fields!gap_value = txtGapVal.Text
            .Fields!remarks = txtRemarks.Text
        Case 2
            .Fields!sched_type = cboSchedules.Text
            .Fields!last_delivery = Format(txtLast.Text, "mm/dd/yyyy")
        Case 3
            .Fields!Description = txtSchedTitle.Text
            .Fields!sched_date = Format(txtSchedDate.Text, "mm/dd/yyyy")
            .Fields!remarks = txtSchedRmrks.Text
    End Select
    .Update
End With
MsgBox msg, vbInformation
ClrFlds
End Sub

Private Sub cmdViewSup_Click()
If txtSupplier.Text = Empty Then
    Exit Sub
End If
Screen.MousePointer = 11
With frmSuppliers
    .ExecSrch "company", txtSupplier.Text
    .cmdSave.Enabled = False
    .cmdNew.Enabled = False
    .cmdDelete.Enabled = False
    .cmdEdit.Enabled = False
    .Show 1
End With
End Sub

Private Sub dtpLastDate_Change()
txtLast.Text = Format(dtpLastDate.Value, "mm/dd/yyyy")
End Sub

Private Sub dtpSchedDate_Change()
txtSchedDate.Text = Format(dtpSchedDate.Value, "mm/dd/yyyy")
mvwSched.Value = dtpSchedDate.Value
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetLv lstRecords, True, True
For i = 2 To freScheduler.UBound
    freScheduler(1).Height = tabScheduler.Height - 420
    freScheduler(1).Width = tabScheduler.Width - 110
    freScheduler(1).Top = tabScheduler.Top + 340
    freScheduler(1).Left = tabScheduler.Left + 30
    freScheduler(i).Move _
        freScheduler(1).Left, _
        freScheduler(1).Top, _
        freScheduler(1).Width, _
        freScheduler(1).Height
    freScheduler(i).Visible = False
Next i
tabScheduler.SelectedItem = tabScheduler.Tabs(1)
freScheduler(1).Visible = True
SetUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
mdiMain.cmdNotifications.Caption = Notifications & " Notifications"
Screen.MousePointer = 0
End Sub

Private Sub lstRecords_DblClick()
If NoRcrd(lstRecords, "No available record on the list.") = True Then Exit Sub
loadData tabScheduler.SelectedItem.Key, lstRecords.ColumnHeaders(2).Text, lstRecords.SelectedItem.SubItems(1), tabScheduler.SelectedItem.Index
End Sub

Public Sub loadData(Table As String, RcrdFld As String, RcrdStr As String, TabIndex As Integer)
If NoRcrd(lstRecords) = True Then Exit Sub
RunSql "Select * from " & Table & " where " & RcrdFld & " = '" & RcrdStr & "'"
With Rs
    Select Case TabIndex
        Case 1
            If .Fields!Description = "(none)" Then
                MsgBox "You cannot update this record. It is a system record", vbCritical
                Exit Sub
            End If
            txtDescription.Text = .Fields!Description
            cboGap.Text = .Fields!Gap
            txtGapVal.Text = .Fields!gap_value
            txtRemarks.Text = .Fields!remarks
        Case 2
            txtSupplier.Text = .Fields!company
            cboSchedules.Text = .Fields!sched_type
            txtLast.Text = Format(.Fields!last_delivery, "mm/dd/yyyy")
            lblPersonel.Caption = .Fields!p_name
        Case 3
            txtSchedTitle.Text = .Fields!Description
            txtSchedDate.Text = Format(.Fields!sched_date, "mm/dd/yyyy")
            txtSchedRmrks.Text = .Fields!remarks
            mvwSched.Value = txtSchedDate.Text
    End Select
End With
cmdSave.Caption = "&Update"
End Sub

Private Sub tabScheduler_Click()
For i = 1 To tabScheduler.Tabs.Count
    If freScheduler(i).Index = tabScheduler.SelectedItem.Index Then
        freScheduler(i).Visible = True
    Else
        freScheduler(i).Visible = False
    End If
Next i
RunSql "Select * from " & tabScheduler.SelectedItem.Key
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.ListIndex = 0
ExecSrch tabScheduler.SelectedItem.Key, "record_no", "%"
ClrFlds
End Sub

Private Sub txtLast_GotFocus()
SelAll txtLast
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch tabScheduler.SelectedItem.Key, cboFilter.Text, txtSrchStr.Text
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStr_GotFocus()
If txtSrchStr = "Search" Then
    txtSrchStr.Text = Empty
    txtSrchStr.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStr_LostFocus()
If Trim(txtSrchStr) = Empty Then
    txtSrchStr.Text = "Search"
    txtSrchStr.ForeColor = &H8000000B
End If
End Sub

Public Sub ExecSrch(Table As String, RcrdFld As String, RcrdStr As String)
RunSql "Select * from " & Table & " where " & RcrdFld & " LIKE '" & RcrdStr & "%' order by record_no ASC"
With Rs
    n = 0
    lstRecords.ColumnHeaders.Clear
    For i = 1 To (.Fields.Count)
        lstRecords.ColumnHeaders.Add
        If n < .Fields.Count Then
            lstRecords.ColumnHeaders(i).Text = .Fields(n).Name
            If lstRecords.ColumnHeaders(i).Text = "remarks" Then
                lstRecords.ColumnHeaders(i).Width = 3000
            End If
        End If
        n = n + 1
    Next i
    lstRecords.ColumnHeaders(1).Text = "#"
    lstRecords.ColumnHeaders(1).Width = 150
    lstRecords.ListItems.Clear
    While Not .EOF = True
        Set x = lstRecords.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub SetUp()
cboGap.ListIndex = 0
DtpValue dtpSchedDate
DtpValue dtpLastDate
mvwSched.Value = Now
mvwSched.Left = freScheduler(3).Width - mvwSched.Width
mvwSched.Top = freScheduler(3).Height - mvwSched.Height
End Sub

Private Sub ClrFlds()
txtDescription.Text = Empty
cboGap.ListIndex = 0
cboGap.ListIndex = 0
txtGapVal.Text = Empty
txtRemarks.Text = Empty

lblPersonel.Caption = "---"
txtSupplier.Text = Empty
LoadCbo "tblDeliverySched", cboSchedules, "description"
cboSchedules.ListIndex = 0
txtLast.Text = "  /  /    "

txtSchedDate.Text = "  /  /    "
txtSchedTitle.Text = Empty
txtSchedRmrks.Text = Empty
ExecSrch tabScheduler.SelectedItem.Key, "record_no", "%"
SetUp
cmdSave.Caption = "&Save"
End Sub
