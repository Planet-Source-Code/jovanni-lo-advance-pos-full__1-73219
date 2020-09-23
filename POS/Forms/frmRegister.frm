VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSrchStrInven 
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
      TabIndex        =   20
      Text            =   "Search"
      Top             =   6000
      Width           =   2895
   End
   Begin VB.ComboBox cboFilterInven 
      Height          =   315
      ItemData        =   "frmRegister.frx":038A
      Left            =   120
      List            =   "frmRegister.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   6000
      Width           =   1575
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
      TabIndex        =   6
      Top             =   960
      Width           =   5295
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmRegister.frx":038E
         Left            =   240
         List            =   "frmRegister.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   9
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
         TabIndex        =   8
         Text            =   "Search"
         Top             =   240
         Width           =   2655
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4800
         Picture         =   "frmRegister.frx":0392
         Top             =   120
         Width           =   480
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
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5295
      Begin VB.ComboBox cboCondition 
         Height          =   315
         ItemData        =   "frmRegister.frx":0C5C
         Left            =   1560
         List            =   "frmRegister.frx":0C5E
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1320
         Width           =   3615
      End
      Begin ctrlButton.ThemedButton cmdViewCondition 
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   840
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
         MouseIcon       =   "frmRegister.frx":0C60
         Picture         =   "frmRegister.frx":0E3A
         PictureAlign    =   2
         PictureSize     =   0
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label2 
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
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblPrice 
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
         Left            =   4680
         TabIndex        =   17
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label5 
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
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label4 
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
         TabIndex        =   14
         Top             =   1320
         Width           =   1005
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   645
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1931
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
         Text            =   "#"
         Object.Width           =   265
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P-Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selling Price"
         Object.Width           =   1764
      EndProperty
   End
   Begin ComctlLib.ListView lvwInventory 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   5295
      _ExtentX        =   9340
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
      NumItems        =   5
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
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selling Price"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "QTY"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   1764
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   1800
      TabIndex        =   21
      Top             =   6120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdRemove 
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      MouseIcon       =   "frmRegister.frx":118E
      Picture         =   "frmRegister.frx":1368
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdView 
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Inventory <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":16BC
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   1560
      TabIndex        =   26
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
      MouseIcon       =   "frmRegister.frx":1896
      Picture         =   "frmRegister.frx":1A70
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   5400
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
      MouseIcon       =   "frmRegister.frx":1DC4
      Picture         =   "frmRegister.frx":1F9E
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdReg 
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Register"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":22F2
      Picture         =   "frmRegister.frx":24CC
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4920
      Picture         =   "frmRegister.frx":2820
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Remove to unregister an item from list"
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
      TabIndex        =   12
      Top             =   8400
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmRegister.frx":30EA
      Top             =   8280
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of unregistered Items on Database."
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
      TabIndex        =   10
      Top             =   3840
      Width           =   3360
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmRegister.frx":39B4
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTER ITEMS"
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
      Width           =   1530
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register an item to Inventory list"
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
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000016&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClrFlds()
lblPcode.Caption = "---"
lblPrice.Caption = "---"
txtDescription.Text = Empty
ViewUnreg "p_code", "%"
ViewInventory "p_code", "%"
End Sub

Private Sub cboFilter_Click()
txtSrchStr_Change
End Sub

Private Sub cboFilterInven_Click()
txtSrchStrInven_Change
End Sub

Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdLoad_Click()
lvwItems_DblClick
End Sub

Private Sub cmdReg_Click()
If lblPcode.Caption = "---" Then MsgBox "Please load an item from the list.", vbExclamation: Exit Sub
If CboEmp(cboCondition) = True Then Exit Sub
RunSql "Select * from tblInventory"
With Rs
    SubSql "Select * from tblItems where p_code = '" & lblPcode.Caption & "'"
    SubRs.Fields!on_inventory = 1
    SubRs.Update
    .AddNew
    .Fields!p_code = SubRs.Fields!p_code
    .Fields!Description = SubRs.Fields!Description
    .Fields!brand_name = SubRs.Fields!brand_name
    .Fields!unit_price = SubRs.Fields!unit_price
    .Fields!net_price = SubRs.Fields!net_price
    .Fields!bar_code = SubRs.Fields!bar_code
    .Fields!Supplier = SubRs.Fields!Supplier
    .Fields!sold = 0
    .Fields!Condition = cboCondition.Text
    .Fields!location = SubRs.Fields!location
    .Fields!date_added = Format(Date, "mm/dd/yyyy")
    SubSql "Select * from tblUnregStocks where p_code = '" & lblPcode.Caption & "'"
    If SubRs.EOF = False Then
        .Fields!quantity = SubRs.Fields!stored_qty
        SubRs.Delete
    Else
        .Fields!quantity = 0
    End If
    .Update
End With
frmInventory.ViewInven "p_code", "%"
frmInventory.ViewItems "p_code", "%"
MsgBox "Item " & lblPcode.Caption & " has been successfully registered to Inventory.", vbInformation
ClrFlds
End Sub

Public Sub ViewInventory(RcrdFld As String, RcrdStr As String)
RunSql "Select p_code, description, net_price, quantity, location from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
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

Private Sub cmdRemove_Click()
ExecRemove lvwInventory.SelectedItem
End Sub

Public Sub ExecRemove(Pcode As String)
If NoRcrd(lvwInventory, "No record available on the list. Please Search for an item.") = True Then Exit Sub
If MsgBox("Your about to remove item " & Pcode & " from your inventory list." & _
            " Removing this item may affect system report." & vbNewLine & vbNewLine & _
            "Do you want to continue?", vbExclamation + vbOKCancel) = vbOK Then
    s = MsgBox("Would you like to save the inventory quantity of this item?", vbExclamation + vbYesNoCancel)
    If s <> vbCancel Then
        If s = vbYes Then
            RunSql "Select * from tblUnregStocks"
            With Rs
                .AddNew
                .Fields!p_code = Pcode
                .Fields!stored_qty = Val(lvwInventory.SelectedItem.SubItems(3))
                .Update
            End With
        End If
        RunSql "Delete * from tblInventory where p_code = '" & Pcode & "'"
        RunSql "Select on_inventory from tblItems where p_code = '" & Pcode & "'"
        Rs.Fields!on_inventory = 0
        Rs.Update
        MsgBox "Item " & Pcode & " has been removed from inventory.", vbInformation
    End If
End If
ClrFlds
frmInventory.ViewInven "p_code", "%"
frmInventory.ViewItems "p_code", "%"
End Sub

Private Sub cmdview_Click()
If cmdView.Caption = "&Inventory >>" Then
    Me.Height = 9255
    cmdView.Caption = "&Inventory <<"
Else
    Me.Height = 6375
    cmdView.Caption = "&Inventory >>"
End If
End Sub

Private Sub cmdViewCondition_Click()
If CboEmp(cboCondition) = True Then Exit Sub
frmItemSettings.loadData "tblStatus", "description", cboCondition.Text
frmItemSettings.Show 1
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
RunSql "Select p_no, p_code, description, net_price from tblItems"
With Rs
    cboFilter.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.Text = "description"

RunSql "Select p_code, description, net_price, quantity, location from tblInventory"
With Rs
    cboFilterInven.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilterInven.AddItem (.Fields(i).Name)
    Next i
End With
cboFilterInven.Text = "description"

SetLv lvwItems, True, True
SetLv lvwInventory, True, True
ViewUnreg "p_code", "%"
ViewInventory "p_code", "%"
LoadCbo "tblStatus", cboCondition, "description", "Select", 1
End Sub

Public Sub ViewUnreg(RcrdFld As String, RcrdStr As String)
RunSql "Select p_no, p_code, description, net_price from tblItems where " & RcrdFld & " LIKE '" & RcrdStr & "%' and on_inventory = 0 Order By p_no ASC"
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

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
mdiMain.cmdWarnings.Caption = Warnings & " Warnings"
Screen.MousePointer = 0
End Sub

Public Sub lvwItems_DblClick()
If NoRcrd(lvwItems, "No record available on the list. Please Search for an item.") = True Then Exit Sub
RunSql "Select p_no, p_code, description, net_price from tblItems where p_code = '" & lvwItems.SelectedItem.SubItems(1) & "'"
With Rs
    lblPcode.Caption = .Fields!p_code
    txtDescription.Text = .Fields!Description
    lblPrice.Caption = "P " & Format(.Fields!net_price, "#0.00")
End With
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ViewUnreg cboFilter.Text, txtSrchStr.Text
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

Private Sub txtSrchStrInven_Change()
If Right(txtSrchStrInven.Text, 1) = "'" Then
    txtSrchStrInven.Text = Empty
End If
If Trim(txtSrchStrInven.Text) <> Empty Then
    If txtSrchStrInven.Text <> "Search" Then
        ViewInventory cboFilterInven.Text, txtSrchStrInven.Text
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStrInven_GotFocus()
If txtSrchStrInven = "Search" Then
    txtSrchStrInven.Text = Empty
    txtSrchStrInven.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStrInven_LostFocus()
If Trim(txtSrchStrInven) = Empty Then
    txtSrchStrInven.Text = "Search"
    txtSrchStrInven.ForeColor = &H80000011
End If
End Sub
