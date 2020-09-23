VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmStocks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   9015
   ClientLeft      =   5685
   ClientTop       =   2355
   ClientWidth     =   8190
   Icon            =   "frmStocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Item Details"
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
      TabIndex        =   25
      Top             =   2760
      Width           =   7935
      Begin VB.Label lblExpiry 
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
         Left            =   1440
         TabIndex        =   34
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblSupplier 
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
         Left            =   3600
         TabIndex        =   31
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier:"
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
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Expiry:"
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
         Index           =   16
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblOnHand 
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
         Left            =   7320
         TabIndex        =   27
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "On Hand:"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaction Details"
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
      TabIndex        =   18
      Top             =   1920
      Width           =   7935
      Begin VB.Label lblOnInven 
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
         Left            =   7320
         TabIndex        =   33
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "On Inventory:"
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
         Left            =   5880
         TabIndex        =   32
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
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
         Left            =   2640
         TabIndex        =   23
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblDate 
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
         Left            =   3600
         TabIndex        =   22
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Items:"
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
         TabIndex        =   21
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblItemCount 
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
         Left            =   1440
         TabIndex        =   20
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Invoice Number"
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
      TabIndex        =   8
      Top             =   960
      Width           =   7935
      Begin VB.ComboBox cboLot 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1575
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
         Left            =   6720
         TabIndex        =   17
         Top             =   360
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
         Left            =   4800
         TabIndex        =   16
         Top             =   360
         Width           =   180
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
         Left            =   2640
         TabIndex        =   15
         Top             =   360
         Width           =   180
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
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   480
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
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   570
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
         Left            =   6120
         TabIndex        =   12
         Top             =   360
         Width           =   450
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
         Left            =   5880
         TabIndex        =   11
         Top             =   360
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
         Index           =   11
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   105
      End
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
      Left            =   4800
      TabIndex        =   6
      Text            =   "Search"
      Top             =   6450
      Width           =   2895
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      ItemData        =   "frmStocks.frx":038A
      Left            =   2280
      List            =   "frmStocks.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6435
      Width           =   2175
   End
   Begin ComctlLib.ListView lvwInventory 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2778
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
      NumItems        =   8
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
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selling Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "QTY"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Discount"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   53
   End
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   4560
      TabIndex        =   5
      Top             =   6600
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lvwList 
      Height          =   1455
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2566
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
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   6840
      TabIndex        =   35
      Top             =   5880
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
      MouseIcon       =   "frmStocks.frx":038E
      Picture         =   "frmStocks.frx":0568
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   5520
      TabIndex        =   36
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Options <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStocks.frx":08BC
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   4200
      TabIndex        =   37
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cl&ear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStocks.frx":0A96
      Picture         =   "frmStocks.frx":0C70
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdPay 
      Height          =   375
      Left            =   2880
      TabIndex        =   38
      Top             =   5880
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
      MouseIcon       =   "frmStocks.frx":0FC4
      Picture         =   "frmStocks.frx":119E
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdProcess 
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   39
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Pr&ocess"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStocks.frx":14F2
      Picture         =   "frmStocks.frx":16CC
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDiscount 
      Height          =   375
      Left            =   6840
      TabIndex        =   40
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Discount"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStocks.frx":1A20
      Picture         =   "frmStocks.frx":1BFA
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdView 
      Height          =   375
      Left            =   5520
      TabIndex        =   41
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&View"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStocks.frx":1F4E
      Picture         =   "frmStocks.frx":2128
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDeduct 
      Height          =   375
      Left            =   4200
      TabIndex        =   42
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Deduct"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStocks.frx":247C
      Picture         =   "frmStocks.frx":2656
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdAdd 
      Height          =   375
      Left            =   2880
      TabIndex        =   43
      Top             =   8520
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
      MouseIcon       =   "frmStocks.frx":29AA
      Picture         =   "frmStocks.frx":2B84
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image imgPay 
      Height          =   480
      Left            =   7560
      Picture         =   "frmStocks.frx":2ED8
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Select an item for task"
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
      TabIndex        =   24
      Top             =   8640
      Width           =   1605
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmStocks.frx":3B1C
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3480
      Picture         =   "frmStocks.frx":43E6
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Click Process to add stocks to inventory from delivery list."
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
      Left            =   3960
      TabIndex        =   19
      Top             =   3840
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Delivery List"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7680
      Picture         =   "frmStocks.frx":4CB0
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmStocks.frx":557A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCKS"
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
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add and deduct stocks from inventory"
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
      TabIndex        =   0
      Top             =   480
      Width           =   2370
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000016&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Total As Double, Payed As Double, Loan As Double

Private Sub cboLot_Click()
If cboLot.ListIndex = 0 Then ClrFlds: Exit Sub
loadData cboLot.Text
End Sub

Private Sub loadData(Lot As String)

RunSql "Select * from tblDeliveryTrans where lot_num = '" & Lot & "'"
With Rs
    lblTotal.Caption = "P " & Format(.Fields!Total, "#0.00")
    lblPayed.Caption = "P " & Format(.Fields!Payed, "#0.00")
    lblLoan.Caption = "P " & Format(.Fields!Loan, "#0.00")
    lblDate.Caption = Format(.Fields!tran_date, "mm/dd/yyyy")
    lblSupplier.Caption = .Fields!Supplier
    ExecSrch "supplier", lblSupplier.Caption
End With
RunSql "Select * from tblStockList where lot_num = '" & Lot & "'"
With Rs
    lblItemCount.Caption = .RecordCount
    lvwList.ListItems.Clear
    While Not .EOF = True
        Set x = lvwList.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 4)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
lvwList_Click
End Sub

Private Sub cmdAdd_Click()
If NoRcrd(lvwInventory, "No available items on the list. Please search for items.") = True Then Exit Sub
If MsgBox("Your about to add stocks on item " & lvwInventory.SelectedItem & ". This is a manual adding of stocks to your inventory. " & vbNewLine & _
            "The transaction will not be included on Delivery Transactions." & vbNewLine & _
        vbNewLine & "Do want to continue?", vbExclamation + vbOKCancel) = vbCancel Then
        Exit Sub
End If
i = ValBox("Input quantity to add.", imgIcon, , , "add quantity")
Screen.MousePointer = 11
RunSql "Select * from tblInventory where p_code = '" & lvwInventory.SelectedItem & "'"
With Rs
    n = .Fields!quantity + i
    .Fields!quantity = n
    .Update
End With
MsgBox "Added " & i & " on item " & lvwInventory.SelectedItem & ". This item has a total quantity of " & n & " on your inventory.", vbInformation
ExecSrch "p_code", lvwInventory.SelectedItem
frmInventory.ViewInven "p_code", "%"
Screen.MousePointer = 0
End Sub

Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDeduct_Click()
If NoRcrd(lvwInventory, "No available record on your inventory. Please search for an item.") = True Then Exit Sub
If MsgBox("Are you sure you want to deduct the quantity of this item?", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If
i = ValBox("Input quantity to deduct", imgIcon, App.Title, , "STOCKS")
RunSql "Select quantity from tblInventory where p_code = '" & lvwInventory.SelectedItem & "'"
With Rs
    If i > .Fields!quantity Then
        MsgBox "The quantity inputed is to much. You only have " & .Fields!quantity & " left on your inventory.", vbExclamation
        Exit Sub
    End If
    .Fields!quantity = .Fields!quantity - i
    .Update
End With
MsgBox "Deducted " & i & " on item " & lvwInventory.SelectedItem & ".", vbInformation
ExecSrch "p_code", lvwInventory.SelectedItem
frmInventory.ViewInven "p_code", "%"
End Sub

Private Sub cmdDiscount_Click()
If NoRcrd(lvwInventory, "No available record on the list. Please search for an item.") = True Then Exit Sub
RunSql "Select discount from tblInventory where p_code = '" & lvwInventory.SelectedItem & "'"
With Rs
    n = ValBox("Input value (e.g. 0.1 = 10%)", imgIcon, App.Title, .Fields!discount, "STOCKS")
    If n < 0 Then Exit Sub
    If n > 1 Then MsgBox "Invalid discount value. Discount must not be greater than 1(100%).", vbExclamation: Exit Sub
    .Fields!discount = n
    .Update
End With
MsgBox "Discount of item " & lvwInventory.SelectedItem & " has been changed to " & n * 100 & "%", vbInformation
ExecSrch "p_code", lvwInventory.SelectedItem
frmInventory.ViewInven "p_code", "%"
End Sub

Private Sub cmdOptions_Click()
If cmdOptions.Caption = "&Options >>" Then
    Me.Height = 9495
    cmdOptions.Caption = "&Options <<"
Else
    Me.Height = 6855
    cmdOptions.Caption = "&Options >>"
End If
End Sub

Private Sub cmdPay_Click()
If CboEmp(cboLot) = True Then Exit Sub
RunSql "Select * from tblDeliveryTrans where lot_num = '" & cboLot.Text & "'"
With Rs
    n = ValBox("Input amount to pay", imgPay, App.Title, .Fields!Loan, "pay loan")
    Payed = .Fields!Payed + n
    Loan = .Fields!Total - Payed
    .Fields!Payed = Payed
    .Fields!Loan = Loan
    .Update
    lblPayed.Caption = Format(Payed, "#0.00")
    lblLoan.Caption = Format(Loan, "#0.00")
End With
MsgBox "Delivery transaction has been updated.", vbInformation
End Sub

Private Sub cmdProcess_Click()
If NoRcrd(lvwList, "No record available on the list. Please select a lot number to process.") = True Then Exit Sub
n = 0
For i = 1 To lvwList.ListItems.Count
    Screen.MousePointer = 11
    RunSql "Select * from tblInventory where p_code = '" & lvwList.ListItems(i).SubItems(1) & "'"
    With Rs
        If .EOF = False Then
            SubSql "Select on_inventory from tblStockList where lot_num = '" & cboLot.Text & "' and p_code = '" & lvwList.ListItems(i).SubItems(1) & "'"
            If SubRs.Fields!on_inventory = 1 Then
                Screen.MousePointer = 0
                MsgBox "Item " & lvwList.ListItems(i).SubItems(1) & "'s stock are already added on inventory.", vbExclamation
            Else
                n = n + 1
                .Fields!quantity = .Fields!quantity + Val(lvwList.ListItems(i).SubItems(3))
                .Update
                SubRs.Fields(0) = 1
                SubRs.Update
            End If
        Else
            Screen.MousePointer = 0
            MsgBox "Item " & lvwList.ListItems(i).SubItems(1) & " is not yet registered on your inventory.", vbExclamation
        End If
    End With
Next i
cboLot.ListIndex = 0
ExecSrch "supplier", lblSupplier.Caption
MsgBox n & " item(s) were updated on your inventory list.", vbInformation

frmInventory.ViewInven "p_code", "%"
Screen.MousePointer = 0
End Sub

Private Sub cmdView_Click()
If NoRcrd(lvwInventory, "No available record on the list. Please search for an item.") = True Then Exit Sub
Screen.MousePointer = 11
frmAddItem.cmdSave.Enabled = False
frmAddItem.cmdNew.Enabled = False
frmAddItem.ExecSrch "p_code", lvwInventory.SelectedItem
frmAddItem.Show 1
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetLv lvwInventory, True, True
SetLv lvwList, True, True
LoadCbo "tblDeliveryTrans", cboLot, "lot_num", "Select", 1
RunSql "Select p_code, description, supplier, unit_price, " & _
        " net_price, quantity, discount, location  from tblInventory"
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.Text = "p_code"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
mdiMain.cmdWarnings.Caption = Warnings & " Warnings"
Screen.MousePointer = 0
End Sub

Private Sub lvwInventory_DblClick()
cmdView_Click
End Sub

Private Sub lvwList_Click()
If NoRcrd(lvwList) = True Then Exit Sub
lblOnHand.Caption = lvwList.SelectedItem.SubItems(3)
lblExpiry.Caption = lvwList.SelectedItem.SubItems(5)
RunSql "Select * from tblInventory where p_code = '" & lvwList.SelectedItem.SubItems(1) & "'"
With Rs
    If .EOF = False Then
        lblOnInven.Caption = .Fields!quantity
    Else
        lblOnInven.Caption = 0
    End If
End With
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

Public Sub ExecSrch(RcrdFld As String, RcrdStr As String)
RunSql "Select p_code, description, supplier, unit_price, " & _
        " net_price, quantity, discount, location  from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
With Rs
    lvwInventory.ListItems.Clear
    While Not .EOF = True
        Set x = lvwInventory.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
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

Private Sub ClrFlds()
lblTotal.Caption = "---"
lblPayed.Caption = "---"
lblLoan.Caption = "---"
lblItemCount.Caption = "---"
lblDate.Caption = "---"
lblSupplier.Caption = "---"
lblOnHand.Caption = "---"
lblOnInven.Caption = "---"
cboLot.ListIndex = 0
lvwList.ListItems.Clear
ExecSrch "p_code", "%"
lblExpiry.Caption = "---"
End Sub
