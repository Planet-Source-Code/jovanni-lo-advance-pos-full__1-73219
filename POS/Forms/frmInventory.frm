VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmInventory 
   Caption         =   "Inventory"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   Icon            =   "frmInventory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   7620
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Edit"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delete"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Search"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Register"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Close"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInventory.frx":038A
      Left            =   720
      List            =   "frmInventory.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInventory.frx":038E
      Left            =   3240
      List            =   "frmInventory.frx":03B9
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin ComctlLib.ListView lvwInventory 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
      _ExtentX        =   4260
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P-Code"
         Object.Width           =   1235
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
         Text            =   "QTY"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Brand Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Unit Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selling Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Bar Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sold"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Condition"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Discount"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   12
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Added"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
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
      NumItems        =   18
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
         Object.Width           =   1411
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
         Text            =   "Brand Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Supplier Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Unit Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selling Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "VAT"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Product ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   12
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Bar Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   13
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Supplier"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   14
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Usage"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(16) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   15
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reg"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(17) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   16
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Image"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   17
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Added"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
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
      TabIndex        =   7
      Top             =   840
      Width           =   435
   End
   Begin VB.Label lblInventory 
      AutoSize        =   -1  'True
      Caption         =   "Inventory - List of registered items on inventory"
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
      TabIndex        =   3
      Top             =   3360
      Width           =   4110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Items - List of items on Database"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   2820
   End
   Begin ComctlLib.ImageList lstMenu 
      Left            =   6840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":0434
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":0786
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":0AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":0E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":117C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":14CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":1820
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInventory.frx":1B72
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboMonth_Click()
ViewItems "p_code", "%"
End Sub

Private Sub cboYear_Click()
ViewItems "p_code", "%"
End Sub

Private Sub Form_Load()
SetLv lvwInventory, True, True
SetLv lvwItems, True, True
mdiMain.tbrMenu.Buttons(2).Value = tbrPressed
mdiMain.sbrStatus.Panels(5).Text = "Click Search for more search options    F3 - Lock the computer"
For i = Format(Date, "yyyy") To 2000 Step -1
    cboYear.AddItem i
Next i
cboYear.AddItem "(View All)"
cboYear.Text = Val(Format(Date, "yyyy"))
cboMonth.ListIndex = Val(Format(Date, "mm"))
n = 1
With tbrMenu
    .ImageList = lstMenu
    For i = 1 To lstMenu.ListImages.Count
        .Buttons(i).Image = i
        n = n + 2
    Next i
End With
ViewItems "p_code", "%"
ViewInven "p_code", "%"
End Sub

Public Sub ViewItems(RcrdFld As String, RcrdStr As String)
If cboYear.Text = "(View All)" Or cboMonth.Text = "(none)" Then
    RunSql "Select * from tblItems Order by p_no ASC"
Else
    RunSql "Select * from tblItems where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(reg_date, 'm') = " & _
            cboMonth.ListIndex & " and format(reg_date,'yyyy') = " & cboYear.Text & " Order by p_no ASC"
End If
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
If NoRcrd(lvwItems) = True Then Exit Sub
lvwItems.ListItems(1).Selected = True
End Sub
Public Sub ViewInven(RcrdFld As String, RcrdStr As String)
RunSql "Select * from tblInventory where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
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
If NoRcrd(lvwInventory) = True Then Exit Sub
lvwInventory.ListItems(1).Selected = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
lvwItems.Width = ScaleWidth - (lvwItems.Left + 100)
lvwItems.Height = (ScaleHeight - lvwItems.Top) / 2.5
lblInventory.Top = lvwItems.Height + lvwItems.Top + lvwItems.Left
lvwInventory.Top = lblInventory.Top + (lblInventory.Height * 2)
lvwInventory.Width = ScaleWidth - (lvwItems.Left + 100)
lvwInventory.Height = ScaleHeight - (lvwInventory.Top + 100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.tbrMenu.Buttons(2).Value = tbrUnpressed
mdiMain.sbrStatus.Panels(5).Text = Empty
End Sub

Private Sub lvwInventory_DblClick()
If NoRcrd(lvwInventory, "No items on your inventory list. Please search for an item.") = True Then Exit Sub
frmStocks.ExecSrch "p_code", lvwInventory.SelectedItem
frmStocks.Show 1
End Sub

Private Sub lvwInventory_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mdiMain.mnuRegister
End If
End Sub

Private Sub lvwItems_DblClick()
If NoRcrd(lvwItems, "No items on your item list. Please search for an item.") = True Then Exit Sub
frmAddItem.ExecSrch "p_code", lvwItems.SelectedItem.SubItems(1)
frmAddItem.Show 1
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mdiMain.mnuInventory
End If
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As ComctlLib.Button)
ExecButtons Button.Index
End Sub
Public Sub ExecButtons(Index As Integer)
Screen.MousePointer = 11
Select Case Index
    Case 1
        frmAddItem.Show 1
    Case 2
        lvwItems_DblClick
    Case 3
        If NoRcrd(lvwItems, "No items on your item list. Please search for an item.") = True Then Exit Sub
        RunSql "Select * from tblInventory where p_code = '" & lvwItems.SelectedItem.SubItems(1) & "'"
        With Rs
            Screen.MousePointer = 0
            If .EOF = False Then
                x = MsgBox("This item, " & lvwItems.SelectedItem.SubItems(1) & _
                            ", is registered on your inventory. Removing this item will affect your inventory records. " & _
                            vbNewLine & vbNewLine & "Do you want to continue?", vbExclamation + vbYesNo)
                If x = vbYes Then
                    SubSql "select * from tblItems where p_code = '" & lvwItems.SelectedItem.SubItems(1) & "'"
                    If SubRs.Fields!image_name <> Empty Then
                        Kill App.Path & "\Images\Products\" & SubRs.Fields!image_name
                    End If
                    SubRs.Delete
                    MsgBox "Item " & lvwItems.SelectedItem.SubItems(1) & " has been deleted", vbInformation
                End If
            Else
                x = MsgBox("Your about to delete item " & lvwItems.SelectedItem.SubItems(1) & ", are you sure?", vbExclamation + vbYesNo)
                If x = vbYes Then
                    SubSql "Select * from tblItems where p_code = '" & lvwItems.SelectedItem.SubItems(1) & "'"
                    If SubRs.Fields!image_name <> Empty Then
                        Kill App.Path & "\Images\Products\" & SubRs.Fields!image_name
                    End If
                    SubRs.Delete
                End If
            End If
                
        End With
        ViewItems "p_code", "%"
        ViewInven "p_code", "%"
    Case 4
        frmItemSettings.Show 1
    Case 5
        frmSrchOpt.Show 1
    Case 6
        If NoRcrd(lvwItems, "No available record from the list. Please search for a record.") = True Then Exit Sub
        RunSql "Select * from tblItems where p_code = '" & lvwItems.SelectedItem.SubItems(1) & "' and on_inventory = 0"
        With Rs
            If .EOF = False Then
                frmRegister.ViewUnreg "p_code", lvwItems.SelectedItem.SubItems(1)
                frmRegister.lvwItems_DblClick
            End If
        End With
        frmRegister.Show 1
    Case 7
        frmStatus.ExecSrch "p_code", lvwInventory.SelectedItem
        frmStatus.cmdLoad_Click
        frmStatus.Show 1
    Case 8
        Unload Me
End Select
Screen.MousePointer = 0
End Sub
