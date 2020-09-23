VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Date Added"
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
      Left            =   4560
      TabIndex        =   16
      Top             =   960
      Width           =   2415
      Begin VB.Label lblDateAdded 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   225
      End
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   6855
      Begin VB.TextBox txtRemarks 
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1800
         Width           =   4935
      End
      Begin VB.ComboBox cboCondition 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         ItemData        =   "frmStatus.frx":038A
         Left            =   4800
         List            =   "frmStatus.frx":038C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPcode 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin ctrlButton.ThemedButton cmdViewSup 
         Height          =   375
         Left            =   6360
         TabIndex        =   34
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
         MouseIcon       =   "frmStatus.frx":038E
         Picture         =   "frmStatus.frx":0568
         PictureAlign    =   2
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdViewItem 
         Height          =   375
         Left            =   3240
         TabIndex        =   31
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
         MouseIcon       =   "frmStatus.frx":08BC
         Picture         =   "frmStatus.frx":0A96
         PictureAlign    =   2
         PictureSize     =   0
      End
      Begin VB.Label lblLocation 
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
         TabIndex        =   24
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   3840
         TabIndex        =   23
         Top             =   1320
         Width           =   765
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
         Left            =   360
         TabIndex        =   20
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Condition:"
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
         TabIndex        =   18
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblBrand 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Brand Name:"
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
         TabIndex        =   14
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblQty 
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
         TabIndex        =   13
         Top             =   840
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
         Index           =   2
         Left            =   3840
         TabIndex        =   12
         Top             =   840
         Width           =   750
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
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item Code:"
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
         Top             =   360
         Width           =   930
      End
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
      Width           =   4335
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmStatus.frx":0DEA
         Left            =   240
         List            =   "frmStatus.frx":0DEC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1215
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
         Left            =   1800
         TabIndex        =   5
         Text            =   "Search"
         Top             =   240
         Width           =   1935
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3720
         Picture         =   "frmStatus.frx":0DEE
         Top             =   120
         Width           =   480
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lstInventory 
      Height          =   1695
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   6855
      _ExtentX        =   12091
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
      NumItems        =   7
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
         Text            =   "QTY"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Brand Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Condition"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Added"
         Object.Width           =   2540
      EndProperty
   End
   Begin ctrlButton.ThemedButton cmdView 
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&View <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStatus.frx":16B8
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   4680
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
      MouseIcon       =   "frmStatus.frx":1892
      Picture         =   "frmStatus.frx":1A6C
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdUpdate 
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   30
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Update"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStatus.frx":1DC0
      Picture         =   "frmStatus.frx":1F9A
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5760
      TabIndex        =   32
      Top             =   7440
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
      MouseIcon       =   "frmStatus.frx":22EE
      Picture         =   "frmStatus.frx":24C8
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   4440
      TabIndex        =   33
      Top             =   7440
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
      MouseIcon       =   "frmStatus.frx":281C
      Picture         =   "frmStatus.frx":29F6
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click an item from the list to load details."
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
      TabIndex        =   27
      Top             =   7560
      Width           =   3405
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "frmStatus.frx":2D4A
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of registered items on Inventory"
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
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmStatus.frx":3614
      Top             =   4680
      Width           =   480
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
      TabIndex        =   22
      Top             =   4815
      Width           =   1845
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update the product status"
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
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM STATUS"
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
      Width           =   1230
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmStatus.frx":3EDE
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdLoad_Click()
If NoRcrd(lstInventory, "No record available on your inventory list. Please search for an item.") = True Then Exit Sub
RunSql "Select p_code, quantity, brand_name, supplier, condition, location, date_added  from tblInventory where p_code = '" & lstInventory.SelectedItem & "'"
With Rs
    txtPcode.Text = .Fields!p_code
    lblBrand.Caption = .Fields!brand_name
    cboCondition.Text = .Fields!Condition
    cboSupplier.Text = .Fields!Supplier
    lblQty.Caption = .Fields!quantity
    lblLocation.Caption = .Fields!location
    lblDateAdded.Caption = Format(.Fields!date_added, "mm/dd/yyyy")
    SubSql "Select * from tblItemCondition where p_code = '" & lstInventory.SelectedItem & "'"
    With SubRs
        If .EOF = False Then
            txtRemarks.Text = .Fields!remarks
        Else
            txtRemarks.Text = Empty
        End If
    End With
End With
End Sub

Private Sub cmdUpdate_Click()
If txtPcode.Text = Empty Then
    MsgBox "No espicified item. Click Load to load an item from your inventory list", vbExclamation
    Exit Sub
End If
RunSql "Select condition from tblInventory where p_code = '" & txtPcode.Text & "'"
With Rs
    .Fields!Condition = cboCondition.Text
    .Update
End With
SubSql "Select * from tblItemCondition where p_code = '" & txtPcode.Text & "'"
With SubRs
    If .EOF = True Then
        If Trim(txtRemarks.Text) <> Empty Then
            .AddNew
            .Fields!record_no = RcrdId("tblItemCondition", , "record_no")
            .Fields!p_code = txtPcode.Text
            .Fields!Condition = cboCondition.Text
            .Fields!remarks = txtRemarks.Text
            .Fields!date_updated = Format(Date, "mm/dd/yyyy")
            .Update
        End If
    Else
        .Fields!remarks = txtRemarks.Text
        .Fields!date_updated = Format(Date, "mm/dd/yyyy")
        .Update
    End If
End With
MsgBox "Item " & txtPcode.Text & "'s condition has been updated.", vbInformation
ExecSrch "p_code", "%"
frmInventory.ViewInven "p_code", "%"
End Sub

Private Sub cmdview_Click()
If cmdView.Caption = "&View <<" Then
    Me.Height = 5595
    cmdView.Caption = "&View >>"
Else
    Me.Height = 8355
    cmdView.Caption = "&View <<"
End If
End Sub

Private Sub cmdViewItem_Click()
If txtPcode.Text = Empty Then Exit Sub
Screen.MousePointer = 11
frmAddItem.cmdSave.Enabled = False
frmAddItem.cmdNew.Enabled = False
frmAddItem.ExecSrch "p_code", txtPcode.Text
frmAddItem.Show 1
End Sub

Private Sub cmdViewSup_Click()
If cboSupplier.Text = Empty Then Exit Sub
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

Public Sub ExecSrch(RcrdFld As String, RcrdStr As String)
RunSql "Select p_code, quantity, brand_name, supplier, condition, location, date_added  from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
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
SetLv lstInventory, True, True
RunSql "Select p_code, quantity, brand_name, supplier, condition, location  from tblInventory"
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.Text = "p_code"

LoadCbo "tblStatus", cboCondition, "description"
LoadCbo "tblSuppliers", cboSupplier, "company"

ExecSrch "p_code", "%"
End Sub

Private Sub lstInventory_DblClick()
cmdLoad_Click
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

Private Sub ClrFlds()
txtPcode.Text = Empty
lblBrand.Caption = "---"
cboCondition.ListIndex = 0
txtRemarks.Text = Empty
cboSupplier.ListIndex = 0
lblQty.Caption = "---"
lblLocation.Caption = "---"
lblDateAdded.Caption = "---"
ExecSrch "p_code", "%"
End Sub
