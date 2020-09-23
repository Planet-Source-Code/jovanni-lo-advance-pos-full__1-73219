VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "TEST"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "&TEST"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdTest_Click()
    'MsgBox Format(Now, "mm/dd/yyyy/h/n/s")
   ' You must close the recordset before changing the parameter.
    'Set x = dtaGroups.rsItemSummary
    'CloseRs x
        PathToDoc = App.Path & "\Run.bat"
        ShellExecute 1, "open", PathToDoc, vbNullString, vbNullString, 5
End Sub
