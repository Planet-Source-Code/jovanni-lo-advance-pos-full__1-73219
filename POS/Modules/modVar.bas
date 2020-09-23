Attribute VB_Name = "modVar"
Public FrstUsr As Boolean
Public UserLvl As String, _
        UserNme As String, _
        UserId As String, _
        PcId As String, _
        UserNo As Integer
Public DetectionType As Integer
Global n As Double, i As Long, x As Variant, s As String, d As Date
Public msg As String
Public ImgName As String, ImgSrc As String

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Sub Main()
    InitCommonControls
    frmxSplash.Show
End Sub

