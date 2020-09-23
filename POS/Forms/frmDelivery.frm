VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   Icon            =   "frmDelivery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin CtrlLine.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   4200
      TabIndex        =   27
      Top             =   3480
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
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
      Left            =   4440
      TabIndex        =   25
      Text            =   "Search"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      ItemData        =   "frmDelivery.frx":038A
      Left            =   2520
      List            =   "frmDelivery.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame Frame3 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   6495
      Begin MSComCtl2.DTPicker dtpExpiry 
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
         Left            =   6000
         TabIndex        =   31
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20709379
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtExpiry 
         Height          =   285
         Left            =   4920
         TabIndex        =   32
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblPcode 
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
         Left            =   4920
         TabIndex        =   29
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P-Code:"
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
         Left            =   3720
         TabIndex        =   28
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date:"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Price:"
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
         TabIndex        =   21
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblSupPrice 
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
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblDescription 
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
         Height          =   435
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "* Invoice Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   16
      Top             =   960
      Width           =   3255
      Begin VB.TextBox txtLot 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   3135
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         ItemData        =   "frmDelivery.frx":038E
         Left            =   180
         List            =   "frmDelivery.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin ctrlButton.ThemedButton cmdViewSup 
         Height          =   375
         Left            =   2640
         TabIndex        =   39
         Top             =   360
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
         MouseIcon       =   "frmDelivery.frx":0392
         Picture         =   "frmDelivery.frx":056C
         PictureAlign    =   2
         PictureSize     =   0
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lvwList 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "#"
         Object.Width           =   265
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P-Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qty"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Amount"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Expiry Date"
         Object.Width           =   1764
      EndProperty
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P-Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Supplier Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selling Price"
         Object.Width           =   1764
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   3360
      TabIndex        =   26
      Top             =   3600
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Height          =   375
      Left            =   5400
      TabIndex        =   35
      Top             =   7680
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
      MouseIcon       =   "frmDelivery.frx":08C0
      Picture         =   "frmDelivery.frx":0A9A
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   4080
      TabIndex        =   36
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDelivery.frx":0DEE
      Picture         =   "frmDelivery.frx":0FC8
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5400
      TabIndex        =   37
      Top             =   5400
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
      MouseIcon       =   "frmDelivery.frx":131C
      Picture         =   "frmDelivery.frx":14F6
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdAdd 
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   38
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDelivery.frx":184A
      Picture         =   "frmDelivery.frx":1A24
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdPay 
      Height          =   375
      Left            =   4080
      TabIndex        =   34
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Pay"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDelivery.frx":1D78
      Picture         =   "frmDelivery.frx":1F52
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
      TabIndex        =   33
      Top             =   5400
      Width           =   1845
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmDelivery.frx":22A6
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   5520
      Picture         =   "frmDelivery.frx":2B70
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPay 
      Height          =   480
      Left            =   4440
      Picture         =   "frmDelivery.frx":37B4
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Click Save to save delivery transaction"
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
      TabIndex        =   30
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmDelivery.frx":43F8
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6240
      Picture         =   "frmDelivery.frx":4CC2
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label lblLvw 
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "|"
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
      Index           =   11
      Left            =   2280
      TabIndex        =   14
      Top             =   5880
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "|"
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
      Index           =   10
      Left            =   4680
      TabIndex        =   13
      Top             =   5880
      Width           =   105
   End
   Begin VB.Label lblLoan 
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
      Left            =   5520
      TabIndex        =   12
      Top             =   5880
      Width           =   180
   End
   Begin VB.Label lblPayed 
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
      Left            =   3240
      TabIndex        =   11
      Top             =   5880
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loan:"
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
      Index           =   9
      Left            =   4920
      TabIndex        =   10
      Top             =   5880
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payed:"
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
      Index           =   8
      Left            =   2520
      TabIndex        =   9
      Top             =   5880
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
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
      TabIndex        =   8
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label lblTotal 
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
      Left            =   960
      TabIndex        =   7
      Top             =   5880
      Width           =   180
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmDelivery.frx":558C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY"
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
      TabIndex        =   6
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View and manage supplier deliveries"
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
      TabIndex        =   5
      Top             =   480
      Width           =   2280
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SupPrice As Double
Dim Total As Double, Payed As Double, Loan As Double
Private Sub cboSupplier_Click()
ExecSrch "p_code", "%"
If cboSupplier.Text = "Select" Then
    lblLvw.Caption = "No supplier selected"
Else
    lblLvw.Caption = "Items from " & cboSupplier.Text
End If
lvwList.ListItems.Clear
lvwItems_Click
End Sub

Private Sub cmdAdd_Click()
If NoRcrd(lvwItems, "No items available on the list. Please select a supplier.") = True Then cboSupplier.SetFocus: Exit Sub
If lblPcode.Caption = "---" Then
    MsgBox "Please select a product from the item list.", vbExclamation
    Exit Sub
End If

If TxtEmp(txtLot) = True Then Exit Sub

n = ValBox("Input desired quantity", imgIcon, App.Title, , "Delivery")
If n = 0 Then Exit Sub
For i = 1 To lvwList.ListItems.Count
    With lvwList.ListItems(i)
        If lvwList.ListItems(i).SubItems(1) = lblPcode.Caption Then
            .SubItems(2) = lblDescription.Caption
            .SubItems(3) = Val(.SubItems(3)) + Format(n, "#0")
            .SubItems(4) = SupPrice * Format(Val(.SubItems(3)), "#0")
            .SubItems(5) = txtExpiry.Text
            ViewTotal
            Exit Sub
        End If
    End With
Next i
Set x = lvwList.ListItems.Add(, , lvwList.ListItems.Count + 1)
x.SubItems(1) = lblPcode.Caption
x.SubItems(2) = lblDescription.Caption
x.SubItems(3) = Format(n, "#0")
x.SubItems(4) = SupPrice * Format(n, "#0")
x.SubItems(5) = txtExpiry.Text
ViewTotal
txtExpiry.Text = "  /  /    "
End Sub

Private Sub ViewTotal()
Total = 0
For i = 1 To lvwList.ListItems.Count
    Total = Total + Val(lvwList.ListItems(i).SubItems(4))
Next i
lblTotal.Caption = "P " & Format(Total, "#,##0.00")

Loan = Total - Payed
lblPayed.Caption = "P " & Format(Payed, "#,##0.00")
lblLoan.Caption = "P " & Format(Loan, "#,##0.00")
End Sub

Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdLoad_Click()
lvwItems_Click
End Sub

Private Sub cmdPay_Click()
If NoRcrd(lvwList, "No items on your delivery list.") = True Then Exit Sub
n = ValBox("Input amount to pay", imgPay, App.Title, Total, "pay transaction")
If n > Total Then
    MsgBox "Your payed amount is greater than the total amount. The system will only get " & Total & ".", vbInformation
    n = Total
End If
Payed = n
ViewTotal
End Sub

Private Sub cmdSave_Click()
If NoRcrd(lvwList, "No items on the list to save.") = True Then Exit Sub
RunSql "Select * from tblDeliveryTrans where lot_num = '" & txtLot.Text & "'"
With Rs
    If .EOF = False Then
        MsgBox "The Invoice Number is already on database. The system cannot process this transaction." & vbNewLine & vbNewLine & _
                "Error saving transaction.", vbCritical
        ClrFlds
        Exit Sub
    End If
    .AddNew
    .Fields!record_no = RcrdId("tblDeliveryTrans", , "record_no")
    .Fields!lot_num = txtLot.Text
    .Fields!Supplier = cboSupplier.Text
    .Fields!Total = Total
    .Fields!Payed = Payed
    .Fields!Loan = Loan
    .Fields!tran_date = Format(Date, "mm/dd/yyyy")
    .Update
End With

RunSql "Select * from tblStockList"
With Rs
    For i = 1 To lvwList.ListItems.Count
        .AddNew
        .Fields(0) = RcrdId("tblStockList", , "record_no")
        For n = 1 To (.Fields.Count - 4)
            .Fields(n) = lvwList.ListItems(i).SubItems(n)
        Next n
        .Fields!lot_num = txtLot.Text
        .Fields!on_inventory = 0
        .Fields!date_added = Format(Date, "mm/dd/yyyy")
        .Update
    Next i
End With

RunSql "Select * from tblSuppliers where company = '" & cboSupplier.Text & "'"
With Rs
    .Fields!last_delivery = Date
    .Update
End With

MsgBox "Delivery transaction has been successfully saved.", vbInformation
ClrFlds
End Sub

Private Sub cmdViewSup_Click()
If cboSupplier.ListIndex = 0 Then Exit Sub
Screen.MousePointer = 11
With frmSuppliers
    .ExecSrch "company", cboSupplier.Text
    .cmdSave.Enabled = False
    .cmdNew.Enabled = False
    .cmdDelete.Enabled = False
    .cmdEdit.Enabled = False
    .Show 1
End With
End Sub

Private Sub dtpExpiry_Change()
txtExpiry.Text = Format(dtpExpiry.Value, "mm/dd/yyyy")
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetLv lvwItems, True, True
SetLv lvwList, True, True
DtpValue dtpExpiry
LoadCbo "tblSuppliers", cboSupplier, "company", "Select", 1

RunSql "Select p_code, description, supplier_price, net_price from tblItems"
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.Text = "description"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If NoRcrd(lvwList) = True Then Exit Sub
If MsgBox("Are you sure you want to close this transaction?", vbExclamation + vbOKCancel) = vbCancel Then
    Cancel = 1
End If
End Sub

Private Sub lvwItems_Click()
If NoRcrd(lvwItems) = True Then Exit Sub
RunSql "Select * from tblItems where p_code = '" & lvwItems.SelectedItem & "'"
With Rs
    lblDescription.Caption = .Fields!Description
    SupPrice = .Fields!supplier_price
    lblSupPrice.Caption = "P " & Format(SupPrice, "#0.00")
    lblPcode.Caption = .Fields!p_code
End With
End Sub

Private Sub lvwItems_DblClick()
cmdAdd_Click
End Sub

Private Sub ClrFlds()
txtLot.Text = Empty
lblDescription.Caption = "---"
lblPcode.Caption = "---"
lblSupPrice.Caption = "---"
lvwList.ListItems.Clear
lblTotal.Caption = "---"
lblPayed.Caption = "---"
lblLoan.Caption = "---"
ExecSrch "p_code", "%"
End Sub

Public Sub ExecSrch(RcrdFld As String, RcrdStr As String)
RunSql "Select p_code, description, supplier_price, net_price from tblItems where " & RcrdFld & " LIKE '" & RcrdStr & "%' and supplier = '" & cboSupplier.Text & "'"
With Rs
    lvwItems.ListItems.Clear
    While Not .EOF = True
        Set x = lvwItems.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub txtExpiry_GotFocus()
SelAll txtExpiry
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
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

