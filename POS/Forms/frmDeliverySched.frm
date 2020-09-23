VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliverySched 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suplliers"
   ClientHeight    =   4320
   ClientLeft      =   825
   ClientTop       =   5430
   ClientWidth     =   7455
   Icon            =   "frmDeliverySched.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7455
   Begin VB.Frame freCalendar 
      Caption         =   "Calendar - [now]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   3135
      Begin MSComCtl2.MonthView mvwDelivery 
         Height          =   2310
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483644
         Appearance      =   0
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
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   20709377
         TitleBackColor  =   16744448
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483632
         CurrentDate     =   40099
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   53
   End
   Begin VB.Frame Frame1 
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
      Height          =   2655
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   3975
      Begin ctrlButton.ThemedButton ThemedButton1 
         Height          =   375
         Left            =   3120
         TabIndex        =   21
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
         MouseIcon       =   "frmDeliverySched.frx":038A
         Picture         =   "frmDeliverySched.frx":0564
         PictureAlign    =   2
         PictureSize     =   0
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         ItemData        =   "frmDeliverySched.frx":08B8
         Left            =   1320
         List            =   "frmDeliverySched.frx":08BA
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   2450
         Picture         =   "frmDeliverySched.frx":08BC
         Top             =   1770
         Visible         =   0   'False
         Width           =   240
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
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSched 
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
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Schedule:"
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
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label lblNext 
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
         Left            =   1320
         TabIndex        =   11
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label lblLast 
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
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Last:"
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
         TabIndex        =   9
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lblLastAdded 
         AutoSize        =   -1  'True
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
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
         Left            =   1440
         TabIndex        =   8
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Next:"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblDelivery 
         AutoSize        =   -1  'True
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
         Left            =   1440
         TabIndex        =   6
         Top             =   960
         Width           =   45
      End
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
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
      MouseIcon       =   "frmDeliverySched.frx":0C46
      Picture         =   "frmDeliverySched.frx":0E20
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdScheduler 
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   3840
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
      MouseIcon       =   "frmDeliverySched.frx":1174
      Picture         =   "frmDeliverySched.frx":134E
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "frmDeliverySched.frx":16A2
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a supplier to view schedule"
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
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View delivery schedule"
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
      Width           =   1455
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY SCHEDULE"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmDeliverySched.frx":1F6C
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   -360
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmDeliverySched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub loadData(RcrdStr As String)
Dim sDate As Date
RunSql "Select * from tblSuppliers where company = '" & RcrdStr & "'"
If Rs.Fields!sched_type = "(none)" Then
    ClrFlds
    lblSched.Caption = Rs.Fields!sched_type
    imgWarning.Visible = False
    Exit Sub
End If
SubSql "Select * from tblDeliverySched where description = '" & Rs.Fields!sched_type & "'"
With SubRs
    sDate = Rs.Fields!last_delivery
    lblSched.Caption = .Fields!Description
    lblLast.Caption = Format(sDate, "mmm. dd, yyyy")
    lblNext.Caption = Format(Scheduler(Format(sDate, "mm"), _
                        Format(sDate, "dd"), _
                        Format(sDate, "yyyy"), _
                        .Fields!gap_value, _
                        .Fields!Gap), "mmm. dd, yyyy")
    txtRemarks.Text = .Fields!remarks
    mvwDelivery.Value = lblNext.Caption
    freCalendar.Caption = "Calendar - [Next Delivery]"
End With
d = DateValue(lblNext.Caption)
If Format(d, "mm") < Format(Date, "mm") And Format(d, "yyyy") = Format(Date, "yyyy") Then
    imgWarning.Visible = True
Else
    imgWarning.Visible = False
End If
mvwDelivery.Enabled = False
End Sub

Private Sub cboSupplier_Click()
If cboSupplier.ListIndex = 0 Then ClrFlds: Exit Sub
loadData cboSupplier.Text
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdScheduler_Click()
If lblSched.Caption = "---" Then
    frmScheduler.Show 1
Else
    frmScheduler.loadData "tblDeliverySched", "description", lblSched.Caption, 1
    frmScheduler.Show 1
End If
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

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetUp
End Sub

Private Sub SetUp()
With mdiMain.cmdNextDelivery
    Me.Left = .Left
    Me.Top = .Top * 1.45
End With
mvwDelivery.Value = Date
LoadCbo "tblSuppliers", cboSupplier, "company", "(none)", 1
End Sub

Private Sub ClrFlds()
lblSched.Caption = "---"
lblLast.Caption = "---"
lblNext.Caption = "---"
txtRemarks.Text = Empty
mvwDelivery.Value = Date
freCalendar.Caption = "Calendar - [now]"
mvwDelivery.Enabled = True
End Sub


