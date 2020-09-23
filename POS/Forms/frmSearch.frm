VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2595
   ClientLeft      =   7290
   ClientTop       =   4830
   ClientWidth     =   5670
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboSrchBy 
      Height          =   315
      ItemData        =   "frmSearch.frx":038A
      Left            =   3360
      List            =   "frmSearch.frx":0391
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtSrchStr 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
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
      MouseIcon       =   "frmSearch.frx":039D
      Picture         =   "frmSearch.frx":0577
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Options >>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmSearch.frx":08CB
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for a specific item"
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
      TabIndex        =   6
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
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
      TabIndex        =   5
      Top             =   120
      Width           =   765
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearch.frx":0AA5
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
      Width           =   5895
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Click Options to specify search"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   2190
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "frmSearch.frx":16E9
      Top             =   1530
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Search by:"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label lblSrchBy 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Field"
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
      Left            =   1455
      TabIndex        =   1
      Top             =   1080
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmSearch.frx":1FB3
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SrchForm As Variant
Dim SrchTable As String
Private Sub cboSrchBy_Click()
'set the search by the field from the database
    If cboSrchBy.Text <> "Select" Then
        lblSrchBy.Caption = cboSrchBy.Text
    Else
        lblSrchBy.Caption = "p_code"
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOptions_Click()
'hide of view option menus
If cmdOptions.Caption = "&Options >>" Then
    Me.Height = 3075
    cmdOptions.Caption = "&Options <<"
    lblLabel.Caption = "Select a field name from the list"
Else
    Me.Height = 2535
    cmdOptions.Caption = "&Options >>"
    lblLabel.Caption = "Click Options to specify search"
End If
txtSrchStr.SetFocus
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Me.Height = 2535
End Sub

Public Sub Srch(Table As String, Form As Form)
Set SrchForm = Form
SrchTable = Table
RunSql "Select * from " & Table
    With Rs
        For i = 0 To (.Fields.Count - 1)
            cboSrchBy.AddItem (.Fields(i).Name)
        Next i
    End With
    cboSrchBy.ListIndex = 0
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        If SrchTable = "tblInventory" Then
            SrchForm.ViewInven lblSrchBy.Caption, txtSrchStr.Text
        Else
            SrchForm.ViewItems lblSrchBy.Caption, txtSrchStr.Text
        End If
    End If
Else
    If SrchTable = "tblInventory" Then
        SrchForm.ViewInven lblSrchBy.Caption, "%"
    Else
        SrchForm.ViewItems lblSrchBy.Caption, "%"
    End If
End If
End Sub

Private Sub txtSrchStr_DblClick()
SelAll txtSrchStr
End Sub
