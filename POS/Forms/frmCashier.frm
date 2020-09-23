VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCashier 
   Caption         =   "Cashier"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCashier.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "VAT"
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
      Left            =   2520
      TabIndex        =   46
      Top             =   5520
      Width           =   1095
      Begin VB.Label lblVat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "00%"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   705
      End
   End
   Begin ComctlLib.ListView lstOrders 
      Height          =   735
      Left            =   120
      TabIndex        =   44
      Top             =   6840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P-Code"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "QTY"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Gross"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Net"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Brand"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame freTrans 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   6360
      Width           =   4815
      Begin VB.Label lblTranCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "000-000-0000"
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
         Left            =   3480
         TabIndex        =   40
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblSlipLabel 
         AutoSize        =   -1  'True
         Caption         =   "Transaction #:"
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
         Left            =   2040
         TabIndex        =   39
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cashier ID:"
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
         Left            =   0
         TabIndex        =   38
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblCashier 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "00000"
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
         Left            =   1080
         TabIndex        =   37
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label3 
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
         Left            =   1800
         TabIndex        =   36
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   5160
      TabIndex        =   34
      Top             =   120
      Width           =   2175
      Begin ctrlButton.ThemedButton cmdAdd 
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Add to List"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCashier.frx":038A
         Picture         =   "frmCashier.frx":0564
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdSelect 
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Select Item"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCashier.frx":08B8
         Picture         =   "frmCashier.frx":0A92
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdRefresh 
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "&Refresh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCashier.frx":0DE6
         Picture         =   "frmCashier.frx":0FC0
         PictureAlign    =   1
         PictureSize     =   0
      End
   End
   Begin VB.Frame freCash 
      Caption         =   "Gross Amount"
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
      TabIndex        =   30
      Top             =   5520
      Width           =   2295
      Begin VB.Label lblGrossTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "P 00.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   960
         TabIndex        =   31
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame freChange 
      Caption         =   "Discount"
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
      Left            =   3720
      TabIndex        =   29
      Top             =   5520
      Width           =   1335
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "0%"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   4935
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "P 00.00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   2595
         TabIndex        =   28
         Top             =   240
         Width           =   2145
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1935
      Left            =   5160
      TabIndex        =   24
      Top             =   4320
      Width           =   2175
      Begin ctrlButton.ThemedButton cmdRemove 
         Height          =   615
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "&Remove"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCashier.frx":1314
         Picture         =   "frmCashier.frx":14EE
         PictureAlign    =   1
      End
      Begin ctrlButton.ThemedButton cmdProcess 
         Height          =   615
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "&Process"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCashier.frx":2142
         Picture         =   "frmCashier.frx":231C
         PictureAlign    =   1
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3960
      TabIndex        =   14
      Top             =   1920
      Width           =   3375
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Net Price:"
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
         TabIndex        =   33
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label lblNet 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1800
         TabIndex        =   32
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label lblGross 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1800
         TabIndex        =   21
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label12 
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
         Left            =   360
         TabIndex        =   20
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblTaxable 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1800
         TabIndex        =   19
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblQtyInput 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Qty Tendered:"
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
         TabIndex        =   17
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gross Price:"
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
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tax Amount:"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   1080
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3735
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Discount:"
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
         TabIndex        =   42
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblDiscount 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   41
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
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
         TabIndex        =   13
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Qty:"
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
         TabIndex        =   12
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Brand:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblBrand 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   225
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label8 
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
         Left            =   360
         TabIndex        =   8
         Top             =   3960
         Width           =   765
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   225
      End
   End
   Begin VB.Frame freSearch 
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
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmCashier.frx":2F70
         Left            =   240
         List            =   "frmCashier.frx":2F72
         Style           =   2  'Dropdown List
         TabIndex        =   2
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
         TabIndex        =   0
         Text            =   "Search"
         Top             =   240
         Width           =   1215
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image imgSearch 
         Height          =   480
         Left            =   3360
         Picture         =   "frmCashier.frx":2F74
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtDescription 
         Height          =   735
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         TabIndex        =   43
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label16 
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
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblPcode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "---"
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
         Left            =   4395
         TabIndex        =   25
         Top             =   360
         Width           =   225
      End
   End
   Begin ComctlLib.ListView lstInventory 
      Height          =   5295
      Left            =   7440
      TabIndex        =   45
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9340
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P-Code"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Bar Code"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Brand"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgQty 
      Height          =   480
      Left            =   6840
      Picture         =   "frmCashier.frx":383E
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   6120
      Picture         =   "frmCashier.frx":4482
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "List of Selected Items"
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
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   1845
   End
End
Attribute VB_Name = "frmCashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cntr As Integer
Dim ItemPrice As Double, TaxPrice As Double, GrossPrice As Double, NetPrice As Double
Dim NetTotal As Double, GrossTotal As Double
Dim Condition As String
Private Sub cboFilter_Click()
ExecSrch cboFilter.Text, txtSrchStr.Text
End Sub

Private Sub cmdAdd_Click()
If lblPcode.Caption = "---" Then Exit Sub
RunSql "Select * from tblOrders where p_code = '" & lblPcode.Caption & "' and tran_code = '" & lblTranCode.Caption & "'"
With Rs
    If .EOF = True Then
        .AddNew
        .Fields!record_no = Val(RcrdId("tblOrders", , "record_no"))
        .Fields!quantity = Val(lblQtyInput.Caption)
        .Fields!gross_amount = Format(GrossPrice, "#0.00")
        .Fields!net_amount = Format(NetPrice, "#0.00")
        .Fields!vat = Format(TaxPrice, "#0.00")
    Else
        .Fields!quantity = .Fields!quantity + Val(lblQtyInput.Caption)
        .Fields!gross_amount = Format(.Fields!gross_amount + GrossPrice, "#0.00")
        .Fields!net_amount = Format(.Fields!net_amount + NetPrice, "#0.00")
        .Fields!vat = Format(.Fields!vat + TaxPrice, "#0.00")
    End If
    .Fields!tran_code = lblTranCode.Caption
    .Fields!p_code = lblPcode.Caption
    .Fields!Description = txtDescription.Text
    .Fields!cashier_id = lblCashier.Caption
    .Fields!date_sold = Format(Date, "mm/dd/yyyy")
    .Update
End With

RunSql "Select quantity, sold from tblInventory where p_code = '" & lblPcode.Caption & "'"
With Rs
    .Fields!quantity = Val(lblQty.Caption) - Val(lblQtyInput.Caption)
    .Fields!sold = Val(lblQtyInput.Caption)
    .Update
End With
cmdRefresh_Click
End Sub

Private Sub ViewSelected()
RunSql "Select p_code, quantity, gross_amount, net_amount, description from " & _
        "tblOrders where tran_code = '" & lblTranCode.Caption & "' Order By record_no ASC"
With Rs
    lstOrders.ListItems.Clear
    While Not .EOF = True
        Set x = lstOrders.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With

NetTotal = 0
For i = 1 To lstOrders.ListItems.Count
    NetTotal = NetTotal + lstOrders.ListItems(i).SubItems(3)
Next i
n = Format(NetTotal, "#0.00") * (Val(txtDiscount.Text) / 100)
NetTotal = Format(NetTotal, "#0.00") - n
lblTotal.Caption = "P " & Format(NetTotal, "#,##0.00")

GrossTotal = 0
RunSql "Select gross_amount, vat from " & _
        "tblOrders where tran_code = '" & lblTranCode.Caption & "'"
With Rs
    While Not .EOF = True
        GrossTotal = GrossTotal + .Fields!gross_amount
        .MoveNext
    Wend
End With
lblGrossTotal.Caption = "P " & Format(GrossTotal, "#,##0.00")
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    cmdRefresh_Click
End If
End Sub

Public Sub cmdProcess_Click()
If NoRcrd(lstOrders) = True Then cmdRefresh_Click: Exit Sub
n = ValBox("Please enter the cash amount.", imgIcon, , NetTotal, "cashier")
If n = 0 Then Exit Sub
If n < NetTotal Then
    MsgBox "Tendered cash is insufficient. Please try again.", vbExclamation
    cmdProcess_Click
    Exit Sub
End If
frmChange.Change n, NetTotal
frmChange.Show 1
RunSql "Select * from tblTransactions"
With Rs
    .AddNew
    .Fields!tran_no = RcrdId("tblTransactions", , "tran_no")
    .Fields!tran_code = lblTranCode.Caption
    .Fields!cashier_id = UserId
    .Fields!gross_amount = GrossTotal
    .Fields!earned = NetTotal
    .Fields!item_count = lstOrders.ListItems.Count
    .Fields!tran_date = Format(Date, "mm/dd/yyyy")
    .Update
End With
RunSql "Select * from tblOrders where tran_code = '" & lblTranCode.Caption & "'"
With Rs
    While Not .EOF = True
        SubSql "Select * from tblSales"
        SubRs.AddNew
        SubRs.Fields!record_no = RcrdId("tblSales", , "record_no")
        SubRs.Fields!quantity = .Fields!quantity
        SubRs.Fields!gross_amount = .Fields!gross_amount
        SubRs.Fields!net_amount = .Fields!net_amount
        SubRs.Fields!vat = .Fields!vat
        SubRs.Fields!tran_code = .Fields!tran_code
        SubRs.Fields!p_code = .Fields!p_code
        SubRs.Fields!Description = .Fields!Description
        SubRs.Fields!cashier_id = .Fields!cashier_id
        SubRs.Fields!date_sold = Format(Date, "mm/dd/yyyy")
        SubRs.Update
        .MoveNext
    Wend
End With
RunSql "Delete * from tblOrders where tran_code = '" & lblTranCode.Caption & "'"
cmdRefresh_Click
End Sub

Public Sub cmdRefresh_Click()
ClrFlds
txtSrchStr.Text = Empty
txtSrchStr.ForeColor = &H80000008
txtSrchStr.SetFocus
End Sub

Public Sub cmdRemove_Click()
If NoRcrd(lstOrders) = True Then Exit Sub
RunSql "Select quantity, sold from tblInventory where p_code = '" & lstOrders.SelectedItem & "'"
With Rs
    .Fields!quantity = .Fields!quantity + Val(lstOrders.SelectedItem.SubItems(1))
    .Fields!sold = .Fields!sold - Val(lstOrders.SelectedItem.SubItems(1))
    .Update
End With
RunSql "Delete * from tblOrders where p_code = '" & lstOrders.SelectedItem & "' and tran_code = '" & lblTranCode.Caption & "'"
cmdRefresh_Click
End Sub

Public Sub cmdSelect_Click()
If lblPcode.Caption = "---" Then Exit Sub
i = ValBox("Input quantity.", imgQty, , , "cashier")
If i = 0 Then Exit Sub
If i > Val(lblQty.Caption) Then
    MsgBox "You only have " & Val(lblQty.Caption) & " of this item on your inventory. Contact your administrator for this matter." & _
    vbNewLine & vbNewLine & "Please try again.", vbExclamation
    cmdSelect_Click
    Exit Sub
End If
lblQtyInput.Caption = i
n = ItemPrice * (Val(lblDiscount.Caption) / 100)

RunSql "Select * from tblItems where p_code = '" & lblPcode.Caption & "'"
With Rs
    GrossPrice = .Fields!unit_price * Val(lblQtyInput.Caption)
    If .Fields!vat = 1 Then
        SubSql "Select * from tblVat"
        TaxPrice = GrossPrice * SubRs.Fields!Value
    Else
        TaxPrice = 0
    End If
    lblTaxable.Caption = "P " & Format(TaxPrice, "#,##0.00")
    lblGross.Caption = "P " & Format(GrossPrice, "#,##0.00")
    NetPrice = GrossPrice + TaxPrice
    lblNet.Caption = "P " & Format(NetPrice, "#,##0.00")
End With
cmdAdd.SetFocus
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub ExecSrch(RcrdFld As String, RcrdStr As String)
RunSql "Select p_code, description, bar_code, brand_name from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%' and quantity > 0" & Condition
With Rs
    lstInventory.ListItems.Clear
    While Not .EOF = True
        Set x = lstInventory.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub Form_Load()
SetLv lstOrders, True, True
SetLv lstInventory, True, True

Condition = Empty

RunSql "Select description from tblStatus where include = 0"
With Rs
    While Not .EOF = True
        Condition = Condition & " and condition <> '" & .Fields!Description & "'"
        .MoveNext
    Wend
End With

With mdiMain
    .tbrMenu.Buttons(3).Value = tbrPressed
    .sbrStatus.Panels(5).Text = "F5 - Refresh   F6 - Input Discount   F7 - Remove Item   F8 - Process Transaction   F1 - Help"
    .mnuCashier(0).Enabled = True
    .mnuCashier(1).Enabled = True
    .mnuCashier(2).Enabled = True
    .mnuCashier(4).Enabled = True
    .mnuCashier(6).Enabled = True
    .mnuCashier(7).Enabled = True
End With
RunSql "Select p_code, description, bar_code, brand_name from tblInventory"
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.Text = "description"
lblCashier.Caption = UserId
lblTranCode.Caption = RcrdId("tblTransactions", lblCashier.Caption & "-", "tran_no")
ViewSelected
RunSql "Select * from tblVat"
lblVat.Caption = Rs.Fields!Value * 100 & "%"
End Sub

Private Sub Form_Resize()
On Error Resume Next
freSearch.Width = ScaleWidth - (freSearch.Left + 100)
txtSrchStr.Width = freSearch.Width - (txtSrchStr.Left + imgSearch.Width)
imgSearch.Left = txtSrchStr.Width + txtSrchStr.Left
freTrans.Left = ScaleWidth - freTrans.Width
lstOrders.Width = ScaleWidth - (lstOrders.Left + 100)
lstOrders.Height = ScaleHeight - (lstOrders.Top + 100)
lstInventory.Width = ScaleWidth - (lstInventory.Left + 100)
End Sub

Private Sub Form_Unload(Cancel As Integer)

With mdiMain
    .tbrMenu.Buttons(3).Value = tbrUnpressed
    .sbrStatus.Panels(5).Text = Empty
    .mnuCashier(0).Enabled = False
    .mnuCashier(1).Enabled = False
    .mnuCashier(2).Enabled = False
    .mnuCashier(4).Enabled = False
    .mnuCashier(6).Enabled = False
    .mnuCashier(7).Enabled = False
End With
End Sub

Private Sub ExecView(Pcode As String)
If lstInventory.ListItems.Count = 0 Then
    ClrFlds
    Exit Sub
End If
RunSql "Select p_code, description, brand_name, net_price, quantity, discount from tblInventory where p_code = '" & lstInventory.SelectedItem & "'"
With Rs
    lblPcode.Caption = .Fields!p_code
    txtDescription.Text = .Fields!Description
    lblBrand.Caption = .Fields!brand_name
    lblQty.Caption = .Fields!quantity
    ItemPrice = Format(.Fields!net_price, "#0.00")
    lblPrice.Caption = "P " & ItemPrice
    lblDiscount.Caption = (.Fields!discount * 100) & "%"
End With
End Sub

Private Sub lstInventory_Click()
If NoRcrd(lstInventory) = True Then Exit Sub
ExecView lstInventory.SelectedItem
End Sub

Private Sub lstInventory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    Cntr = 0
End If
lstInventory_Click
End Sub

Private Sub lstInventory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSelect_Click
If KeyAscii = vbKeyEscape Then
    txtSrchStr.SetFocus
End If
End Sub

Private Sub lstInventory_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
    If lstInventory.ListItems.Count = 0 Then Exit Sub
    If lstInventory.SelectedItem = lstInventory.ListItems(1) Then
        Cntr = Cntr + 1
        If Cntr = 2 Then
            txtSrchStr.SetFocus
        End If
    Else
        Cntr = 0
    End If
End If
lstInventory_Click
End Sub

Private Sub txtDiscount_GotFocus()
txtDiscount.Text = Val(txtDiscount.Text)
txtDiscount.SelStart = 0
txtDiscount.SelLength = Len(txtDiscount)
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    txtSrchStr.SetFocus
End If
End Sub

Private Sub txtDiscount_LostFocus()
If Val(txtDiscount.Text) > 100 Or Val(txtDiscount.Text) < 0 Then
    txtDiscount.Text = 0
End If
txtDiscount.Text = Val(txtDiscount.Text) & "%"
ViewSelected
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
        lstInventory_Click
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStr_GotFocus()
Cntr = 0
If txtSrchStr = "Search" Then
    txtSrchStr.Text = Empty
    txtSrchStr.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    If Trim(txtSrchStr.Text) = Empty Or NoRcrd(lstInventory) = True Then Exit Sub
    If txtSrchStr.SelStart = Len(txtSrchStr) Then
        lstInventory.SetFocus
        Cntr = 0
        lstInventory.ListItems(1).Selected = True
    End If
End If
End Sub

Private Sub txtSrchStr_LostFocus()
If Trim(txtSrchStr) = Empty Then
    txtSrchStr.Text = "Search"
    txtSrchStr.ForeColor = &H80000011
End If
End Sub

Private Sub ClrFlds()
lblTranCode.Caption = RcrdId("tblTransactions", lblCashier.Caption & "-", "tran_no")
lstInventory.ListItems.Clear
txtDescription.Text = Empty
lblPcode.Caption = "---"
lblBrand.Caption = "---"
lblQty.Caption = "---"
lblPrice.Caption = "---"
lblQtyInput.Caption = "---"
lblGross.Caption = "---"
lblTaxable.Caption = "---"
lblDiscount.Caption = "---"
lblNet.Caption = "---"
ViewSelected
End Sub
