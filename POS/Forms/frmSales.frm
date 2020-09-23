VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSales 
   Caption         =   "Sales"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   9045
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1111
      ButtonWidth     =   1058
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Query"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Search"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Print"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Close"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin CtrlLine.ctrlLiner ctrLine 
      Height          =   30
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboCashier 
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
      ItemData        =   "frmSales.frx":038A
      Left            =   5760
      List            =   "frmSales.frx":0391
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   840
      Width           =   2055
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
      ItemData        =   "frmSales.frx":039F
      Left            =   3240
      List            =   "frmSales.frx":03CA
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   840
      Width           =   1095
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
      ItemData        =   "frmSales.frx":0445
      Left            =   720
      List            =   "frmSales.frx":0447
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin ComctlLib.ListView lvwSales 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3625
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
      NumItems        =   7
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
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Gross Sale"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Net Sale"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "VAT"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "QTY"
         Object.Width           =   529
      EndProperty
   End
   Begin VB.Label Label9 
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
      Left            =   4080
      TabIndex        =   12
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cashier:"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   840
      Width           =   675
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
      TabIndex        =   8
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label8 
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
      TabIndex        =   7
      Top             =   840
      Width           =   585
   End
   Begin VB.Label lblSellable 
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
      Left            =   6360
      TabIndex        =   5
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Most Sellable Item:"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label lblTotalSales 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Sales for the Month:"
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
      Top             =   1560
      Width           =   2175
   End
   Begin ComctlLib.ImageList lstMenu 
      Left            =   240
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSales.frx":0449
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSales.frx":079B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSales.frx":0AED
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSales.frx":0E3F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCashier_Click()
ViewSales "p_code", "%"
End Sub

Private Sub cboMonth_Click()
ViewSales "p_code", "%"
End Sub

Private Sub cboYear_Click()
ViewSales "p_code", "%"
End Sub

Private Sub Form_Load()
For i = Format(Date, "yyyy") To 2000 Step -1
    cboYear.AddItem i
Next i
cboYear.Text = Val(Format(Date, "yyyy"))
cboMonth.ListIndex = Val(Format(Date, "mm"))
SetLv lvwSales, True, True
LoadCbo "tblAccountSecurity", cboCashier, "id", "(Not Specified)", 1
n = 1
With tbrMenu
    .ImageList = lstMenu
    For i = 1 To lstMenu.ListImages.Count
        .Buttons(i).Image = i
        n = n + 2
    Next i
End With
mdiMain.tbrMenu.Buttons(4).Value = tbrPressed
End Sub

Private Sub Form_Resize()
lvwSales.Width = ScaleWidth - (lvwSales.Left + 100)
lvwSales.Height = ScaleHeight - (lvwSales.Top + 100)
ctrLine.Width = ScaleWidth - ctrLine.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.tbrMenu.Buttons(4).Value = tbrUnpressed
End Sub

Public Sub ViewSales(RcrdFld As String, RcrdStr As String)
If cboCashier.ListIndex = 0 Then
    RunSql "Select record_no, p_code, description, gross_amount, net_amount, vat, quantity from tblSales where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(date_sold, 'm') = " & _
            cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text & " Order by record_no ASC"
Else
    RunSql "Select record_no, p_code, description, gross_amount, net_amount, vat, quantity from tblSales where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(date_sold, 'm') = " & _
            cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text & " and cashier_id = '" & cboCashier.Text & "' Order by record_no ASC"
End If
With Rs
    lvwSales.ListItems.Clear
    While Not .EOF = True
        Set x = lvwSales.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
RunSql "Select format(sum(net_amount),'#0.00') from tblSales where format(date_sold, 'm') = " & cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text
lblTotalSales.Caption = "P " & Val(Format(Rs.Fields(0), "#,##0.00"))

RunSql "Select p_code from tblSales where quantity = (Select max(quantity) from tblSales) and format(date_sold, 'm') = " & cboMonth.ListIndex & " and format(date_sold,'yyyy') = " & cboYear.Text
lblSellable.Caption = Rs.Fields(0)
End Sub


