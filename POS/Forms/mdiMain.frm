VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1191
      ButtonWidth     =   1535
      ButtonHeight    =   1032
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Home"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Inventory"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cashier"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sales"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Suppliers"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Accounts"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog dlgRestore 
      Left            =   4680
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8550
      Left            =   0
      ScaleHeight     =   8550
      ScaleWidth      =   1920
      TabIndex        =   2
      Top             =   675
      Width           =   1920
      Begin ctrlButton.ThemedButton cmdCalc 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Calculator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":038A
         Picture         =   "mdiMain.frx":0564
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   4
         Top             =   5880
         Width           =   255
      End
      Begin CtrlLine.ctrlLiner ctrlLiner5 
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   5760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner3 
         Height          =   30
         Left            =   0
         TabIndex        =   5
         Top             =   4920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   0
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner4 
         Height          =   30
         Left            =   -240
         TabIndex        =   8
         Top             =   6600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   53
      End
      Begin ctrlButton.ThemedButton cmdNotes 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Notes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":08B8
         Picture         =   "mdiMain.frx":0A92
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdScheduler 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Scheduler"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":0DE6
         Picture         =   "mdiMain.frx":0FC0
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdStatus 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Status"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1314
         Picture         =   "mdiMain.frx":14EE
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdDelivery 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Delivery"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1842
         Picture         =   "mdiMain.frx":1A1C
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdStocks 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Stoc&ks"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1D70
         Picture         =   "mdiMain.frx":1F4A
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdNextDelivery 
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   4320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Next Delivery"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":229E
         Picture         =   "mdiMain.frx":2478
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdLogout 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Log-in"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":27CC
         Picture         =   "mdiMain.frx":29A6
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdWarnings 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   6840
         Width           =   1455
         _ExtentX        =   2566
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
         MouseIcon       =   "mdiMain.frx":2CFA
         Picture         =   "mdiMain.frx":2ED4
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdNotifications 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   7320
         Width           =   1455
         _ExtentX        =   2566
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
         MouseIcon       =   "mdiMain.frx":3228
         Picture         =   "mdiMain.frx":3402
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quick access"
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
         TabIndex        =   12
         Top             =   480
         Width           =   795
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "mdiMain.frx":3756
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TASK"
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
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SYSTEM"
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
         Index           =   2
         Left            =   720
         TabIndex        =   10
         Top             =   5880
         Width           =   750
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "mdiMain.frx":439A
         Top             =   5880
         Width           =   480
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detections"
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
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   6240
         Width           =   645
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -240
         Top             =   5760
         Width           =   2295
      End
   End
   Begin cPopMenu6.PopMenu cPopMnu 
      Left            =   3240
      Top             =   840
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin VB.Timer tmrDate 
      Interval        =   1000
      Left            =   2760
      Top             =   840
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9225
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5054
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList lstIcons 
      Left            =   3960
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   34
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":4FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":5330
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":5682
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":59D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":5D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":6078
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":63CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":671C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":6A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":6DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":7112
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":7464
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":77B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":7B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":7E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":81AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":84FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":8850
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":8BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":8EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":9246
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":9598
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":98EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":9C3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":9F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":A2E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":A632
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":A984
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":ACD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":B028
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":B37A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":B6CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":BA1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":BD70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList lstMenu 
      Left            =   2160
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":C0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":CD14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":D966
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":E5B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":F20A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":FE5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":10AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":11700
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":12352
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":12FA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Shortcut        =   {F11}
      End
      Begin VB.Menu lSet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "S&ales"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Reports"
         Begin VB.Menu mnuCategory 
            Caption         =   "Cate&gorized"
         End
         Begin VB.Menu mnuAdvance 
            Caption         =   "Ad&vance Query"
         End
         Begin VB.Menu lAdvance 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMonthly 
            Caption         =   "&Sales Reports"
            Begin VB.Menu mnuReportSales 
               Caption         =   "&Montly"
               Index           =   0
            End
            Begin VB.Menu mnuReportSales 
               Caption         =   "&Weekly"
               Index           =   1
            End
            Begin VB.Menu mnuReportSales 
               Caption         =   "&Daily"
               Index           =   2
            End
         End
      End
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Database"
         Begin VB.Menu mnuDb 
            Caption         =   "&Back-up"
            Index           =   0
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuDb 
            Caption         =   "&Restore"
            Index           =   1
         End
         Begin VB.Menu mnuDb 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuDb 
            Caption         =   "&Clear Records"
            Index           =   3
            Begin VB.Menu mnuClr 
               Caption         =   "Inventory"
               Index           =   0
            End
            Begin VB.Menu mnuClr 
               Caption         =   "-"
               Index           =   1
            End
            Begin VB.Menu mnuClr 
               Caption         =   "D&elivery Transactions"
               Index           =   2
            End
            Begin VB.Menu mnuClr 
               Caption         =   "&Cashier Transactions"
               Index           =   3
            End
            Begin VB.Menu mnuClr 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnuClr 
               Caption         =   "&All Records"
               Index           =   5
            End
         End
         Begin VB.Menu mnuDb 
            Caption         =   "&Compact"
            Index           =   4
         End
         Begin VB.Menu mnuDb 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuDb 
            Caption         =   "&Manage"
            Index           =   6
         End
      End
      Begin VB.Menu lSup 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccess 
         Caption         =   "A&ccess"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Loc&k PC"
         Shortcut        =   {F3}
      End
      Begin VB.Menu lLock 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuSuppliers 
      Caption         =   "&Suppliers"
      Begin VB.Menu mnuSupProfile 
         Caption         =   "S&upplier Profiles"
      End
      Begin VB.Menu lSupplier 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScheduler 
         Caption         =   "Manage Sche&duler"
      End
      Begin VB.Menu mnuDelivery 
         Caption         =   "View Deli&very"
      End
   End
   Begin VB.Menu mnuInventoryMain 
      Caption         =   "&Inventory"
      Begin VB.Menu mnuInventory 
         Caption         =   "I&tems on Database"
         Begin VB.Menu mnuInven 
            Caption         =   "&Add"
            Index           =   0
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuInven 
            Caption         =   "&Edit"
            Index           =   1
         End
         Begin VB.Menu mnuInven 
            Caption         =   "&Delete"
            Index           =   2
            Shortcut        =   {DEL}
         End
         Begin VB.Menu mnuInven 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuInven 
            Caption         =   "S&earch"
            Index           =   4
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuInven 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuInven 
            Caption         =   "Se&ttings"
            Index           =   6
         End
         Begin VB.Menu mnuInven 
            Caption         =   "&Register"
            Index           =   7
         End
         Begin VB.Menu mnuInven 
            Caption         =   "&Close"
            Index           =   8
         End
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "In&ventory Items"
         Begin VB.Menu mnuReg 
            Caption         =   "Remo&ve"
            Index           =   0
         End
         Begin VB.Menu mnuReg 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuReg 
            Caption         =   "&Stocks"
            Index           =   2
         End
         Begin VB.Menu mnuReg 
            Caption         =   "Sta&tus"
            Index           =   3
         End
      End
      Begin VB.Menu lReg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Clear Inventory"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "&Accounts"
      Begin VB.Menu mnuUser 
         Caption         =   "&User Profiles"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "&Manage Security"
      End
   End
   Begin VB.Menu mnuCashierMain 
      Caption         =   "Cashier"
      Begin VB.Menu mnuCashier 
         Caption         =   "Se&lect Item"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "Add Item to &List"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "Re&fresh List"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "Input D&iscount"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "Re&move Item"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuCashier 
         Caption         =   "Pr&ocess Transaction"
         Enabled         =   0   'False
         Index           =   7
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuManual 
         Caption         =   "&User's Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu lManual 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Public Sub cmdDelivery_Click()
Screen.MousePointer = 11
frmDelivery.Show 1
End Sub

Private Sub cmdHide_Click()
If cmdHide.Caption = "<<" Then
    picSideBar.Width = 240
    cmdHide.Caption = ">>"
    imgIcon(0).Visible = False
    imgIcon(1).Visible = False
Else
    picSideBar.Width = 1920
    cmdHide.Caption = "<<"
    imgIcon(0).Visible = True
    imgIcon(1).Visible = True
End If
End Sub

Private Sub cmdLogOut_Click()
ClrAccess
frmLogin.Show 1
End Sub

Private Sub ClrAccess()
UserLvl = Empty
UserNme = Empty
UserId = Empty
UserNo = Empty
Unload frmInventory
Unload frmCashier
Unload frmSales
cmdLogout.Caption = "&Log-in"
For i = 1 To 7
    sbrStatus.Panels(i).Text = Empty
Next i
End Sub

Private Sub cmdNextDelivery_Click()
Screen.MousePointer = 11
frmDeliverySched.Show 1
End Sub

Private Sub cmdNotes_Click()
frmNotes.Show 1
End Sub

Private Sub cmdNotifications_Click()
Screen.MousePointer = 11
Notifications 2
frmWarnings.Show 1
End Sub

Public Sub cmdScheduler_Click()
Screen.MousePointer = 11
frmScheduler.Show 1
End Sub

Private Sub cmdStatus_Click()
Screen.MousePointer = 11
frmStatus.Show 1
End Sub

Private Sub cmdStocks_Click()
Screen.MousePointer = 11
frmStocks.Show 1
End Sub

Private Sub cmdWarnings_Click()
Screen.MousePointer = 11
Warnings 1
frmWarnings.Show 1
End Sub

Private Sub MDIForm_Load()
Me.Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
n = 1
With tbrMenu
    .ImageList = lstMenu
    For i = 1 To lstMenu.ListImages.Count
        .Buttons(i).Image = i
        n = n + 2
    Next i
End With

With cPopMnu
    .MenuBackgroundColor = RGB(255, 255, 255)
    .ImageList = lstIcons
    .SubClassMenu Me
    .ItemIcon("mnuSettings") = lstIcons.ListImages(1).Index - 1
    .ItemIcon("mnuSales") = lstIcons.ListImages(2).Index - 1
    .ItemIcon("mnuSupProfile") = lstIcons.ListImages(4).Index - 1
    .ItemIcon("mnuScheduler") = lstIcons.ListImages(5).Index - 1
    .ItemIcon("mnuDelivery") = lstIcons.ListImages(3).Index - 1
    .ItemIcon("mnuAccess") = lstIcons.ListImages(6).Index - 1
    .ItemIcon("mnuLock") = lstIcons.ListImages(28).Index - 1
    .ItemIcon("mnuExit") = lstIcons.ListImages(7).Index - 1
    .ItemIcon("mnuInven(0)") = lstIcons.ListImages(8).Index - 1
    .ItemIcon("mnuInven(1)") = lstIcons.ListImages(9).Index - 1
    .ItemIcon("mnuInven(2)") = lstIcons.ListImages(10).Index - 1
    .ItemIcon("mnuReg(0)") = lstIcons.ListImages(24).Index - 1
    .ItemIcon("mnuInven(7)") = lstIcons.ListImages(26).Index - 1
    .ItemIcon("mnuReg(2)") = lstIcons.ListImages(11).Index - 1
    .ItemIcon("mnuReg(3)") = lstIcons.ListImages(12).Index - 1
    .ItemIcon("mnuRefresh") = lstIcons.ListImages(27).Index - 1
    .ItemIcon("mnuInven(4)") = lstIcons.ListImages(13).Index - 1
    .ItemIcon("mnuInven(6)") = lstIcons.ListImages(14).Index - 1
    .ItemIcon("mnuInven(8)") = lstIcons.ListImages(7).Index - 1
    .ItemIcon("mnuCategory") = lstIcons.ListImages(15).Index - 1
    .ItemIcon("mnuAdvance") = lstIcons.ListImages(16).Index - 1
    .ItemIcon("mnuMonthly") = lstIcons.ListImages(17).Index - 1
    .ItemIcon("mnuUser") = lstIcons.ListImages(18).Index - 1
    .ItemIcon("mnuSecurity") = lstIcons.ListImages(19).Index - 1
    .ItemIcon("mnuManual") = lstIcons.ListImages(20).Index - 1
    .ItemIcon("mnuAbout") = lstIcons.ListImages(21).Index - 1
    .ItemIcon("mnuCashier(0)") = lstIcons.ListImages(12).Index - 1
    .ItemIcon("mnuCashier(1)") = lstIcons.ListImages(9).Index - 1
    .ItemIcon("mnuCashier(2)") = lstIcons.ListImages(22).Index - 1
    .ItemIcon("mnuCashier(4)") = lstIcons.ListImages(25).Index - 1
    .ItemIcon("mnuCashier(6)") = lstIcons.ListImages(24).Index - 1
    .ItemIcon("mnuCashier(7)") = lstIcons.ListImages(23).Index - 1
    .ItemIcon("mnuDb(0)") = lstIcons.ListImages(29).Index - 1
    .ItemIcon("mnuDb(1)") = lstIcons.ListImages(34).Index - 1
    .ItemIcon("mnuDb(3)") = lstIcons.ListImages(30).Index - 1
    .ItemIcon("mnuDb(4)") = lstIcons.ListImages(22).Index - 1
    .ItemIcon("mnuDb(6)") = lstIcons.ListImages(32).Index - 1
    .ItemIcon("mnuDatabase") = lstIcons.ListImages(31).Index - 1
    .ItemIcon("mnuReport") = lstIcons.ListImages(33).Index - 1
End With

sbrStatus.Panels(6).Text = "Today is: " & Format(Now, "dddd, mmm dd, yyyy")
sbrStatus.Panels(7).Text = "Time: " & Format(Time, "hh:mm:ss")

cmdWarnings.Caption = Warnings & " Warnings"
cmdNotifications.Caption = Notifications & " Reminders"

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAccess_Click()
cmdLogOut_Click
End Sub

Private Sub mnuCashier_Click(Index As Integer)
With frmCashier
    Select Case Index
        Case 0
            .cmdSelect_Click
        Case 1
        Case 2
            .cmdRefresh_Click
        Case 4
            .txtDiscount.SetFocus
        Case 6
            .cmdRemove_Click
        Case 7
            .cmdProcess_Click
    End Select
End With
End Sub

Private Sub mnuCategory_Click()
frmRptCategorized.Show 1
End Sub

Private Sub mnuClr_Click(Index As Integer)
Select Case Index
    Case 5
        If MsgBox("WARNING: This will delete all records from your database excluding user accounts. Be sure to backup your current database. " & vbNewLine & vbNewLine & _
                "Are you really sure you want to continue?", vbCritical + vbYesNo) = vbYes Then
            Screen.MousePointer = 11
            RunSql "Delete * from tblItems"
            RunSql "Delete * from tblDeliveryTrans"
            RunSql "Delete * from tblTransactions"
            RunSql "Delete * from tblDeliverySched"
            RunSql "Delete * from tblLocation"
            RunSql "Delete * from tblBrands"
            RunSql "Delete * from tblType"
            RunSql "Delete * from tblSuppliers"
            RunSql "Delete * from tblStatus"
            RunSql "Delete * from tblSchedules"
            Screen.MousePointer = 0
            MsgBox "Your database is now clear. You can now start with a new transactions.", vbInformation
            ClrAccess
        End If
End Select
End Sub

Private Sub mnuDb_Click(Index As Integer)
On Error GoTo ErrMsg
Select Case Index
    Case 0
        If MsgBox("This may take several minutes to complete. Do you want to continue?", vbExclamation + vbOKCancel) = vbOK Then
            Screen.MousePointer = 11
            Con.Close
            MkDir App.Path & "\Backup\" & Format(Now, "mmm-dd-yyyy(hn)")
            FileCopy App.Path & "\POS.mdb", App.Path & "\Backup\" & Format(Now, "mmm-dd-yyyy(hn)") & "\POS.mdb"
            Screen.MousePointer = 0
            MsgBox "Backup has been successfully completed.", vbInformation
            OpenCon
            ClrAccess
        End If
    Case 1
        s = MsgBox("This will replace your current database. Please backup your current database first. " & vbNewLine & vbNewLine & _
                "Do you want to backup it now?", vbExclamation + vbYesNoCancel)
        If s = vbYes Then
            mnuDb_Click 0
        ElseIf s = vbCancel Then
            Exit Sub
        End If
        dlgRestore.DialogTitle = "Restore Database"
        dlgRestore.InitDir = App.Path & "\Backup"
        dlgRestore.Filter = "Access Database(*.mdb)|*.mdb|All Files (*.*)|*.*"
        dlgRestore.ShowOpen
        If dlgRestore.FileName = Empty Or dlgRestore.FileTitle = Empty Then
            Exit Sub
        End If
        Screen.MousePointer = 11
        Con.Close
        FileCopy dlgRestore.FileName, App.Path & "\" & dlgRestore.FileTitle
        Screen.MousePointer = 0
        MsgBox "Database restore has been successfully completed", vbInformation
        OpenCon
        ClrAccess
    Case 3
        
    Case 4
        MsgBox "This may take several minutes to complete.", vbOKOnly + vbInformation
        Screen.MousePointer = 11
        Con.Close
        CompactDB App.Path & "\POS.mdb"
        Screen.MousePointer = 0
        MsgBox "Database compacting has been successfully completed.", vbInformation
        OpenCon
    Case Else
        Exit Sub
End Select
Exit Sub
ErrMsg:
    Screen.MousePointer = 0
    MsgBox ExecErr("An error has accured on system database management." & vbNewLine & vbNewLine & "Error: " & Err.Description), vbCritical
    OpenCon
End Sub

Private Sub mnuDelivery_Click()
mdiMain.cmdDelivery_Click
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuProcess_Click()
frmCashier.cmdProcess_Click
End Sub

Private Sub mnuLock_Click()
frmLock.Show 1
End Sub

Private Sub mnuInven_Click(Index As Integer)
With frmInventory
    Select Case Index
        Case 0
            .ExecButtons 1
        Case 1
            .ExecButtons 2
        Case 2
            .ExecButtons 3
        Case 4
            .ExecButtons 5
        Case 6
            .ExecButtons 4
        Case 7
            .ExecButtons 6
        Case 8
            .ExecButtons 9
    End Select
End With
End Sub

Private Sub mnuManual_Click()
PathToDoc = App.Path & "\help.chm"
ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
End Sub

Private Sub mnuReg_Click(Index As Integer)
With frmInventory
    Select Case Index
        Case 0
            With frmRegister
                If NoRcrd(frmInventory.lvwInventory, "No record available on Inventory. Please search for a record.") = True Then Exit Sub
                .ExecRemove frmInventory.lvwInventory.SelectedItem
            End With
        Case 2
            frmStocks.Show 1
        Case 3
            .ExecButtons 7
    End Select
End With
End Sub

Private Sub mnuSales_Click()
Unload frmCashier
Unload frmInventory
frmSales.Show
End Sub

Private Sub mnuScheduler_Click()
frmScheduler.ExecSrch "tblSuppliers", "record_id", "%"
frmScheduler.tabScheduler.Tabs(2).Selected = True
frmScheduler.Show 1
End Sub

Private Sub mnuSecurity_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
frmAcntManage.Show 1
End Sub

Private Sub mnuSettings_Click()
frmSettings.Show 1
End Sub

Private Sub mnuSupProfile_Click()
frmSuppliers.Show 1
End Sub

Private Sub mnuUser_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
frmAccountProfile.Show 1
End Sub

Private Sub picSideBar_Resize()
cmdHide.Top = picSideBar.Height - cmdHide.Height
cmdHide.Left = picSideBar.Width - cmdHide.Width
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As ComctlLib.Button)
Screen.MousePointer = 11
With tbrMenu
    Select Case Button.Index
        Case 2
            'If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
            Unload frmCashier
            Unload frmSales
            frmInventory.Show
        Case 3
            'If UserLimit(UserLvl, "Cashier") = True Then Exit Sub
            Unload frmInventory
            Unload frmSales
            frmCashier.Show
        Case 4
            mnuSales_Click
        Case 5
            Screen.MousePointer = 0
            PopupMenu mnuReport, x:=Button.Left, y:=.Top + .Height
        Case 6
            Screen.MousePointer = 0
            PopupMenu mnuSuppliers, x:=Button.Left, y:=.Top + .Height
        Case 7
            frmSettings.Show 1
        Case 8
            Screen.MousePointer = 0
            PopupMenu mnuAccounts, x:=Button.Left, y:=.Top + .Height
        Case 9
            Screen.MousePointer = 0
            PopupMenu mnuHelp, x:=Button.Left, y:=.Top + .Height
        Case 10
            End
    End Select
End With
Screen.MousePointer = 0
End Sub

Private Sub tmrDate_Timer()
sbrStatus.Panels(6).Text = "Today is: " & Format(Now, "dddd, mmm dd, yyyy") & "  "
sbrStatus.Panels(7).Text = "Time: " & Format(Time, "hh:mm:ss")
End Sub
