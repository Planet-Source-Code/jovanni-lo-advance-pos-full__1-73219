VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmItemSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmItemSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   13
      Top             =   2400
      Width           =   495
      Begin VB.TextBox txtVatRmrks 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtVat 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         TabIndex        =   33
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
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
         TabIndex        =   32
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   12
      Top             =   2400
      Width           =   615
      Begin VB.TextBox txtTypeRmrks 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtType 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtUnit 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Product Type:"
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
         TabIndex        =   28
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type Unit:"
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
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   2400
      Width           =   375
      Begin VB.TextBox txtStatRmrks 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtStatus 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkInclude 
         Caption         =   "Include a product to inventory with this setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   20
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         TabIndex        =   23
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Include:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1500
         Width           =   675
      End
   End
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   840
      TabIndex        =   10
      Top             =   2400
      Width           =   375
      Begin VB.TextBox txtLocRmrks 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtLocation 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         TabIndex        =   19
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   495
      Begin VB.TextBox txtBrandRmrks 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   525
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtBrand 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Brand name:"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   810
      End
   End
   Begin ComctlLib.TabStrip tabTables 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4471
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Key             =   "tblStatus"
            Object.Tag             =   "tblInventory"
            Object.ToolTipText     =   "condition"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Product Type"
            Key             =   "tblType"
            Object.Tag             =   "tblItems"
            Object.ToolTipText     =   "brand_type"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Brand Names"
            Key             =   "tblBrands"
            Object.Tag             =   "tblItems"
            Object.ToolTipText     =   "brand_name"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Location"
            Key             =   "tblLocation"
            Object.Tag             =   "tblInventory"
            Object.ToolTipText     =   "location"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VAT"
            Key             =   "tblVat"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmItemSettings.frx":038A
         Left            =   240
         List            =   "frmItemSettings.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
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
         Left            =   2160
         TabIndex        =   5
         Text            =   "Search"
         Top             =   240
         Width           =   2895
      End
      Begin CtrlLine.ctrlLiner ctrlLiner3 
         Height          =   30
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5040
         Picture         =   "frmItemSettings.frx":038E
         Top             =   120
         Width           =   480
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lstRecords 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
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
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   41
      Top             =   6600
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
      MouseIcon       =   "frmItemSettings.frx":0C58
      Picture         =   "frmItemSettings.frx":0E32
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   3120
      TabIndex        =   42
      Top             =   6600
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
      MouseIcon       =   "frmItemSettings.frx":1186
      Picture         =   "frmItemSettings.frx":1360
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   4440
      TabIndex        =   43
      Top             =   6600
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
      MouseIcon       =   "frmItemSettings.frx":16B4
      Picture         =   "frmItemSettings.frx":188E
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   1800
      TabIndex        =   44
      Top             =   6600
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
      MouseIcon       =   "frmItemSettings.frx":1BE2
      Picture         =   "frmItemSettings.frx":1DBC
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search or Load a record from the list to update"
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
      TabIndex        =   40
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmItemSettings.frx":2110
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label lblNo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   5520
      TabIndex        =   39
      Top             =   4440
      Width           =   105
   End
   Begin VB.Label lblNoLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Record #:"
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
      Left            =   4560
      TabIndex        =   38
      Top             =   4440
      Width           =   825
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmItemSettings.frx":29DA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SETTINGS"
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
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set inventory record fields"
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
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmItemSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ClrFlds()
txtStatus.Text = Empty
txtStatRmrks.Text = Empty
chkInclude.Value = Empty
txtType.Text = Empty
txtTypeRmrks.Text = Empty
txtUnit.Text = Empty
txtBrand.Text = Empty
txtBrandRmrks.Text = Empty
txtLocation.Text = Empty
txtLocRmrks.Text = Empty
txtVat.Text = Empty
txtVatRmrks.Text = Empty
LoadTable tabTables.SelectedItem.Key, cboFilter.Text, "%"
cmdSave.Caption = "&Save"
lblNo.Caption = Val(RcrdId(tabTables.SelectedItem.Key, , "record_no"))
lblNoLabel.Caption = "Record #(new):"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrMsg
If tabTables.SelectedItem.Index = 5 Then
    MsgBox "Cannot delete the value of VAT. You may change it's value.", vbExclamation
    Exit Sub
End If

If NoRcrd(lstRecords, "No record on the list. Please search for a record and try again.") = True Then Exit Sub
If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo) = vbYes Then
    RunSql "Delete * from " & tabTables.SelectedItem.Key & " where record_no = " & lstRecords.SelectedItem
    MsgBox "Record has been delete", vbInformation
    ClrFlds
End If
Exit Sub
ErrMsg:
    MsgBox ExecErr("Cannot delete this item because it is used by another record.", _
                    "p_code", tabTables.SelectedItem.Tag, _
                    tabTables.SelectedItem.ToolTipText, _
                    lstRecords.SelectedItem.SubItems(1)), vbCritical
End Sub

Private Sub cmdNew_Click()
lstRecords_DblClick
End Sub

Private Sub cmdSave_Click()
RunSql "Select * from " & tabTables.SelectedItem.Key & " where record_no = " & lblNo.Caption
With Rs
    If cmdSave.Caption <> "&Update" Then
        If tabTables.SelectedItem.Index = 5 Then
            MsgBox "You cannot add another value of VAT. You may update it's value.", vbExclamation
            Exit Sub
        End If
        .AddNew
        .Fields!record_no = Val(RcrdId(tabTables.SelectedItem.Key, , "record_no"))
        msg = "Added new record on " & tabTables.SelectedItem.Caption & "."
    Else
        msg = "Record of no" & lblNo.Caption & " from " & tabTables.SelectedItem.Caption & " has been updated."
    End If
    Select Case tabTables.SelectedItem.Index
        Case 1
            If TrapPrimary(tabTables.SelectedItem.Key, "description", txtStatus.Text) = True Then Exit Sub
            .Fields!Description = txtStatus.Text
            .Fields!include = chkInclude.Value
            .Fields!remarks = txtStatRmrks.Text
        Case 2
            If TrapPrimary(tabTables.SelectedItem.Key, "description", txtType.Text) = True Then Exit Sub
            .Fields!Description = txtType.Text
            .Fields!unit = txtUnit.Text
            .Fields!remarks = txtTypeRmrks.Text
        Case 3
            If TrapPrimary(tabTables.SelectedItem.Key, "description", txtBrand.Text) = True Then Exit Sub
            .Fields!Description = txtBrand.Text
            .Fields!remarks = txtBrandRmrks.Text
        Case 4
            If TrapPrimary(tabTables.SelectedItem.Key, "description", txtLocation.Text) = True Then Exit Sub
            .Fields!Description = txtLocation.Text
            .Fields!remarks = txtLocRmrks.Text
        Case 5
            .Fields!Value = Val(txtVat.Text)
            .Fields!remarks = txtVatRmrks.Text
    End Select
    .Update
End With
MsgBox msg, vbInformation
ClrFlds
End Sub

Private Function TrapPrimary(Table As String, RcrdFld As String, RcrdStr As String) As Boolean
If cmdSave.Caption = "&Save" Then
    SubSql "Select * from " & Table & " where " & RcrdFld & " = '" & RcrdStr & "'"
    With SubRs
        If .EOF = False Then
            TrapPrimary = True
            MsgBox "Description of the item is already on database. Please add another", vbExclamation
        Else
            TrapPrimary = False
        End If
    End With
Else
    TrapPrimary = False
End If
End Function

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetLv lstRecords, True, True
For i = 2 To freSettings.UBound
    freSettings(1).Height = tabTables.Height - 420
    freSettings(1).Width = tabTables.Width - 110
    freSettings(1).Top = tabTables.Top + 340
    freSettings(1).Left = tabTables.Left + 30
    freSettings(i).Move _
        freSettings(1).Left, _
        freSettings(1).Top, _
        freSettings(1).Width, _
        freSettings(1).Height
    freSettings(i).Visible = False
Next i
tabTables.SelectedItem = tabTables.Tabs(1)
freSettings(1).Visible = True
End Sub

Public Sub loadData(Table As String, RcrdFld As String, RcrdStr As String)
If NoRcrd(lstRecords, "No record on the list. Please search for a record and try again.") = True Then Exit Sub
If tabTables.SelectedItem.Index = 5 Then RcrdFld = "value"
RunSql "Select * from " & Table & " where " & RcrdFld & " LIKE '" & RcrdStr & "'"
With Rs
    lblNo.Caption = .Fields!record_no
    Select Case tabTables.SelectedItem.Index
        Case 1
            txtStatus.Text = .Fields!Description
            chkInclude.Value = .Fields!include
            txtStatRmrks.Text = .Fields!remarks
        Case 2
            txtType.Text = .Fields!Description
            txtUnit.Text = .Fields!unit
            txtTypeRmrks.Text = .Fields!remarks
        Case 3
            txtBrand.Text = .Fields!Description
            txtBrandRmrks.Text = .Fields!remarks
        Case 4
            txtLocation.Text = .Fields!Description
            txtLocRmrks.Text = .Fields!remarks
        Case 5
            txtVat.Text = .Fields!Value
            txtVatRmrks.Text = .Fields!remarks
    End Select
End With
cmdSave.Caption = "&Update"
lblNoLabel.Caption = "Record #:"
End Sub

Private Sub lstRecords_DblClick()
loadData tabTables.SelectedItem.Key, "description", lstRecords.SelectedItem.SubItems(1)
End Sub

Private Sub tabTables_Click()
For i = 1 To tabTables.Tabs.Count
    If freSettings(i).Index = tabTables.SelectedItem.Index Then
        freSettings(i).Visible = True
    Else
        freSettings(i).Visible = False
    End If
Next i
RunSql "Select * from " & tabTables.SelectedItem.Key
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.ListIndex = 0
ClrFlds
LoadTable tabTables.SelectedItem.Key, cboFilter.Text, "%"
End Sub

Private Sub LoadTable(Table As String, RcrdFld As String, RcrdStr As String)
RunSql "Select * from " & Table & " where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
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

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        LoadTable tabTables.SelectedItem.Key, cboFilter.Text, txtSrchStr.Text
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
    txtSrchStr.ForeColor = &H80000011
End If
End Sub

