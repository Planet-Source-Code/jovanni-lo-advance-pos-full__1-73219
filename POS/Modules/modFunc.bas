Attribute VB_Name = "modFunc"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public PathToDoc As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hWnd As Long, ByVal lpOperation As String, _
                ByVal lpFile As String, ByVal lpParameters As String, _
                ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function NoRcrd(ListView As Variant, Optional Prompt As String) As Boolean
If ListView.ListItems.Count = 0 Then
    Screen.MousePointer = 0
    If Prompt <> Empty Then
        MsgBox Prompt, vbExclamation
    End If
    NoRcrd = True
Else
    NoRcrd = False
End If
End Function

Public Function RcrdId(Table As String, Optional Identifier As String, Optional FldNo As String) As String
Dim RcrdNo As Integer
RunSql "Select * from " & Table & " order by " & FldNo & " ASC"
With Rs
    If Rs.EOF = False Then
        .MoveLast
        RcrdNo = .Fields(FldNo) + 1
    Else
        RcrdNo = 1
    End If
    If Identifier <> Empty Then
        RcrdId = Identifier & RcrdNo & Format(Date, "mm")
    Else
        RcrdId = RcrdNo
    End If
End With
End Function

Public Sub SrchBox(Table As String, Form As Form)
With frmSearch
    .Srch Table, Form
    .Show 1
End With
End Sub

Public Function ValBox(Prompt As String, Icon As Image, Optional Title As String, _
                        Optional Default As Double, _
                        Optional Header As String = "Value Box") As Double
With frmValue
    If Title <> Empty Then
        .Caption = Title
    Else
        .Caption = App.Title
    End If
    .lblHeader.Caption = StrConv(Header, vbUpperCase)
    .imgIcon.Picture = Icon.Picture
    .lblPrompt.Caption = Prompt
    .Default Val(Default)
    .Show 1
    ValBox = Val(.txtValue.Text)
    Unload frmValue
End With
End Function

Public Function StrBox(Prompt As String, Icon As Image, Optional Title As String, _
                        Optional Default As String, _
                        Optional Header As String = "Text Box", _
                        Optional Style As Integer = 0, _
                        Optional AddCbo As Boolean, _
                        Optional Table As String, _
                        Optional RcrdFld As String) As String
With frmString
    If Title <> Empty Then
        .Caption = Title
    Else
        .Caption = App.Title
    End If
    .lblHeader.Caption = StrConv(Header, vbUpperCase)
    .imgIcon.Picture = Icon.Picture
    .lblPrompt.Caption = Prompt
    Select Case Style
        Case 1
            Set x = .txtStr
            x.Text = Default
        Case 2
            Set x = .cboStr
            If AddCbo = True Then
                LoadCbo "tblStatus", .cboStr, "description", Default, 1
            Else
                x.Text = Default
            End If
        Case 3
            Set x = .cboStr2
            If AddCbo = True Then
                LoadCbo "tblStatus", .cboStr2, "description", Default, 1
            Else
                x.Text = Default
            End If
    End Select
    x.TabIndex = 0
    x.Visible = True
    .Show 1
    StrBox = x.Text
    Unload frmString
End With
End Function

Public Sub LoadCbo(Table As String, _
                    cboBox As ComboBox, _
                    FldStr As String, _
                    Optional strDefault As String, _
                    Optional addDefault As Integer, _
                    Optional Locked As Boolean = False)
RunSql "select * from " & Table
With Rs
    cboBox.Clear
    If addDefault = 1 Then
        cboBox.AddItem (strDefault)
    End If
    While Not .EOF = True
        cboBox.AddItem (.Fields(FldStr))
        .MoveNext
    Wend
End With
cboBox.Locked = Locked
If strDefault <> Empty Then
    cboBox.ListIndex = 0
End If
End Sub

Public Function CboEmp(ByRef ComboBox As ComboBox, _
                Optional TabObject As ComctlLib.TabStrip, _
                Optional TabIndex As Integer) _
                As Boolean
'function for combobox value is default (Select)
If ComboBox.ListIndex = 0 Then
    CboEmp = True
    MsgBox "Please select a record from the list.", vbExclamation
    If TabIndex <> Empty Then
        TabObject.SelectedItem = TabObject.Tabs(TabIndex)
    End If
    ComboBox.SetFocus
Else
    CboEmp = False
End If
End Function

Public Sub DtpValue(dtpObject As DTPicker)
dtpObject.Value = Now
End Sub

Public Sub SelAll(ByRef stext As Variant)
'highlight textbox on focus
With stext
.SelStart = 0
.SelLength = Len(stext)
End With
End Sub
Public Function TxtEmp(ByRef stext As Variant, _
                        Optional TabObject As ComctlLib.TabStrip, _
                        Optional TabIndex As Integer) _
                        As Boolean
'if the textbox is empty then TxtEmp = true
If Trim(stext) = Empty Or stext.Text = "  /  /    " Then
    TxtEmp = True
    MsgBox "Please fill in all required fields.", vbExclamation
    If TabIndex <> Empty Then
        TabObject.SelectedItem = TabObject.Tabs(TabIndex)
    End If
    stext.SetFocus
Else
    TxtEmp = False
End If
End Function

Public Function txtNum(ByRef stext As Variant) As Boolean
'if the input is not a numeric then true
If IsNumeric(stext) = False Then
    txtNum = True
    MsgBox "The field requires a numeric value.", vbExclamation
    stext.SetFocus
    SelAll stext
Else
    txtNum = False
End If
End Function

Public Function UserLimit(ByRef lvl As String, ByRef SysLvl As String) As Boolean
'only the administrator can access some stuffs
If lvl = "Administrator" Then
    UserLimit = False
    Exit Function
End If
If SysLvl <> lvl Then
    Screen.MousePointer = 0
    MsgBox "You don`t have the right to access this task. Please Log in as 'Administrator'", vbExclamation
    UserLimit = True
Else
    UserLimit = False
End If
End Function

Public Sub FrmShow(lvl As String)
If lvl = "Administrator" Or lvl = "User" Then
    x = ReadINI(App.Path & "\Settings.ini", "System", "Log-on Task")
    Select Case x
        Case "Inventory"
            frmInventory.Show
        Case "Cashier"
            frmCashier.Show
        Case "Stocks"
            frmStocks.Show 1
        Case "Status"
            frmStatus.Show 1
        Case "Default"
            frmInventory.Show
    End Select
Else
    mdiMain.Show
End If
End Sub

Public Sub SetLv(ListView As ListView, Optional FullRow As Boolean, Optional GridLines As Boolean)
If GridLines = True Then
    LvGrid ListView
End If
If FullRow = True Then
    LvFullRow ListView
End If
End Sub

Public Function ReadINI(strFile As String, strKey As String, strName As String) As String
Dim intLen As Integer
Dim strText As String
strText = "                                                                                                    "
intLen = GetPrivateProfileString(strKey, strName, "", strText, Len(strText), strFile)
If intLen > -1 Then
    strText = Left(strText, intLen)
Else
    MsgBox "Error on reading configuration", vbCritical
    End
End If
ReadINI = strText
End Function

Public Sub WriteINI(strFile As String, strKey As String, strName As String, strText As String)
Dim intLen As Integer
intLen = WritePrivateProfileString(strKey, strName, strText, strFile)
End Sub

Public Function Scheduler(IntM As Integer, _
                        IntD As Integer, _
                        IntY As Integer, _
                        GapVal As Integer, _
                        Optional Gap As String = "Week") As Date
                        
Dim Max As Long, LastVal As Integer
Select Case Gap
    Case "Day"
        Max = Val(ReadINI(App.Path & "\Settings.ini", "Month Max", MonthName(IntM, True)))
        IntD = IntD + GapVal
        For i = 1 To IntD
            If i = Max Then
                LastVal = Max
                IntM = IntM + 1
                If IntM > 12 Then IntM = 1: IntY = IntY + 1
                Max = Max + Val(ReadINI(App.Path & "\Settings.ini", "Month Max", MonthName(IntM, True)))
            End If
        Next i
        IntD = IntD - LastVal
    Case "Week"
        Dim MaxDays As Integer
        MaxDays = 7 * GapVal
        Max = Val(ReadINI(App.Path & "\Settings.ini", "Month Max", MonthName(IntM, True)))
        IntD = IntD + MaxDays
        For i = 1 To IntD
            If i = Max Then
                LastVal = Max
                IntM = IntM + 1
                If IntM > 12 Then IntM = 1: IntY = IntY + 1
                Max = Max + Val(ReadINI(App.Path & "\Settings.ini", "Month Max", MonthName(IntM, True)))
            End If
        Next i
        IntD = IntD - LastVal
    Case "Month"
        IntM = IntM + GapVal
        If IntM > 12 Then IntM = IntM - 12: IntY = IntY + 1
    Case "Year"
        IntY = IntY + GapVal
End Select
Scheduler = DateSerial(IntY, IntM, IntD)
End Function

Public Function Warnings(Optional dType As Integer) As Integer
DetectionType = dType
RunSql "Delete * from tblDetections"

'-------expiration
RunSql "SELECT * From tblStockList"
With Rs
    While Not .EOF = True
        If .Fields!expiry_date <> "  /  /    " Then
            d = Format(DateValue(.Fields!expiry_date), "mm/dd/yyyy")
            If Format(d, "mm") >= Format(Date, "mm") And _
                    Format(d, "mm") <= (Val(Format(Date, "mm")) + 2) And _
                    Format(d, "yyyy") = Format(Date, "yyyy") Then
                SubSql "Select * from tblInventory where p_code = '" & .Fields!p_code & "'"
                If SubRs.EOF = False Then
                    n = SubRs.Fields!quantity
                End If
                If Format(.Fields!expiry_date, "mm") <= Format(Date, "mm") And _
                        Format(.Fields!expiry_date, "dd") <= Format(Date, "dd") And _
                        Format(.Fields!expiry_date, "yyyy") <= Format(Date, "yyyy") Then
                    s = "This item is already expired. Please unregister this from Inventory and add new stocks." & _
                    vbNewLine & vbNewLine & "Item Description: " & .Fields!Description & vbNewLine & vbNewLine & _
                    "Expiry Date: " & Format(.Fields!expiry_date, "mmm. dd, yyyy") & vbNewLine & _
                    "Quantity on Inventory: " & n
                    SaveDetection .Fields!p_code, "Expired", s, "tblDetections"
                Else
                    s = Format(.Fields!expiry_date, "MM") - Format(Date, "MM") _
                        & " Month(s) before Expiry. Please replace it with new stocks and delete your old stocks. " & _
                        vbNewLine & vbNewLine & "Item Description: " & .Fields!Description & vbNewLine & vbNewLine & _
                        "Expiry date: " & Format(.Fields!expiry_date, "mmm. dd, yyyy") & vbNewLine & _
                        "quantity on Inventory: " & n
                    SaveDetection .Fields!p_code, "Expiration", s, "tblDetections"
                End If
            End If
        End If
        .MoveNext
    Wend
End With

'-------out of stock
RunSql "SELECT * From tblInventory WHERE quantity < 10"
With Rs
    While Not .EOF = True
        s = "This item do not have enough quantity on your inventory. Please add stock for this item." & vbNewLine & vbNewLine & _
            "Item Description: " & .Fields!Description & vbNewLine & vbNewLine & _
            "Currently on Inventory: " & .Fields!quantity
        SaveDetection .Fields!p_code, "Low Stock", s, "tblDetections"
        .MoveNext
    Wend
End With

'-------low inventory
RunSql "Select * from tblInventory"
With Rs
    If .RecordCount = 0 Or .RecordCount <= 10 Then
        s = "You don`t have enough items on your inventory." & _
            "Please add items or register items from database to your inventory list." & vbNewLine & vbNewLine & _
            "Items on Inventory: " & .RecordCount
        SaveDetection "Inventory", "Low Inventory", s, "tblDetections"
    End If
End With

'-------no sales for the month
RunSql "Select * from tblInventory"
With Rs
    While Not .EOF = True
        If Format(Date, "mm") <> 1 Then
            n = Format(Date, "mm") - 1
            SubSql "Select * from tblSales where p_code = '" & .Fields!p_code & "' and format([date_sold],'mm') = " & n & _
                "and format([date_sold],'yyyy') = " & Format(Date, "yyyy")
            If SubRs.EOF = False Then
                If SubRs.Fields!quantity < 30 Then
                    i = 0
                    While Not SubRs.EOF = True
                        i = i + SubRs.Fields!quantity
                        SubRs.MoveNext
                    Wend
                    s = "Sales of this item is less for this month." & vbNewLine & vbNewLine & _
                            "Last Month total sales: " & i
                    SaveDetection .Fields!p_code, "Less Sales", s, "tblDetections"
                End If
            End If
        End If
        .MoveNext
    Wend
End With

'-----No supplier
RunSql "Select * from tblSuppliers"
With Rs
    If .RecordCount = 0 Then
        s = "No supplier saved on database. Please add a supplier for item delivery."
        SaveDetection "Suppliers", "No Supplier", s, "tblDetections"
    End If
End With

'-----Items no registered
RunSql "Select * from tblItems where on_inventory = 0"
With Rs
    n = 0
    While Not .EOF = True
        SubSql "SELECT * From tblStockList WHERE p_code = '" & .Fields!p_code & "' and Format$([expiry_date],'mm') Between " _
        & Val(Format(Date, "MM")) & " And " & Val(Format(Date, "MM") + 2) _
        & " and format$(expiry_date, 'yyyy') = " _
        & Format(Date, "yyyy")
        If SubRs.EOF = True Then
            n = n + 1
        End If
        .MoveNext
    Wend
    If n > 0 Then
        s = "Some items on your database are not registered on your inventory list. If you don`t register this items, " & _
            " they will not be included on your sales." & vbNewLine & vbNewLine & _
            "Unregistered Items: " & n
        SaveDetection "Register", "Non-Registered", s, "tblDetections"
    End If
End With

'-----Delivery Schedule exceeded
RunSql "Select Sup.Company as Company, Sup.last_delivery as LastDelivery, sched.gap as Gap, sched.gap_value as GapVal from tblSuppliers as Sup " & _
        "INNER JOIN tblDeliverySched as Sched ON Sup.sched_type = Sched.description"
With Rs
    While Not .EOF = True
        d = Scheduler(Format(.Fields!lastdelivery, "mm"), _
            Format(.Fields!lastdelivery, "dd"), _
            Format(.Fields!lastdelivery, "yyyy"), _
            .Fields!GapVal, _
            .Fields!Gap)
        If Format(d, "mm") <= Format(Date, "mm") And Format(d, "dd") < Format(Date, "dd") And .Fields!Gap <> "(none)" Then
            s = "Delivery schedule of supplier, " & .Fields!company & ", is not updated. " & _
                "Please record all delivery transactions of your suppliers to update it's delivery schedule." & vbNewLine & vbNewLine & _
                "Last Delivery: " & Format(.Fields!lastdelivery, "mmm. dd, yyyy") & vbNewLine & _
                "Expected Date: " & Format(d, "mmm. dd, yyyy")
            SaveDetection .Fields!company, "Delivery Sched", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With
RunSql "Select * from tblDetections"
With Rs
    Warnings = .RecordCount
End With

End Function

Public Function Notifications(Optional dType As Integer) As Integer
DetectionType = dType
RunSql "Delete * from tblDetections"
'-------Delivery
RunSql "Select company, sched_type, last_delivery from tblSuppliers"
With Rs
    While Not .EOF = True
        SubSql "Select * from tblDeliverySched where description = '" & .Fields(1) & "'"
        x = Scheduler(Format(.Fields!last_delivery, "mm"), Format(.Fields!last_delivery, "dd"), Format(.Fields!last_delivery, "yyyy"), SubRs.Fields!gap_value, SubRs.Fields!Gap)
        If Format(Date, "mm") = Format(x, "mm") And SubRs.Fields!Gap <> "(none)" Then
            i = Format(x, "dd") - 7
            If Format(Date, "dd") <= Format(x, "dd") And Format(Date, "dd") >= i Then
                s = Format(x, "dd") - Format(Date, "dd") & " day(s) before delivery of new stocks." & vbNewLine & vbNewLine & _
                    "Company: " & .Fields!company & vbNewLine & vbNewLine & _
                    "Delivery Date: " & Format(x, "mmm. dd, yyyy")
                SaveDetection .Fields!company, "Delivery", s, "tblDetections"
            End If
        End If
        .MoveNext
    Wend
End With

'--------schedule task
RunSql "Select * from tblSchedules"
With Rs
    While Not .EOF = True
        If Format(.Fields!sched_date, "mm/dd/yyyy") = Format(Date, "mm/dd/yyyy") Then
            s = "Scheduled task is today." & vbNewLine & vbNewLine & _
                "Task: " & .Fields!remarks & vbNewLine & vbNewLine & _
                "Date: " & Format(.Fields!sched_date, "mm/dd/yyyy")
            SaveDetection Format(.Fields!Description, "mm/dd/yyyy"), "Scheduled Task", s, "tblDetections"
        End If
        .MoveNext
    Wend
End With

RunSql "Select * from tblDetections"
With Rs
    Notifications = .RecordCount
End With

End Function

Private Sub SaveDetection(Reference As String, Title As String, Description As String, Table As String)
SubSql "Select * from " & Table
With SubRs
    .AddNew
    .Fields!record_no = Val(RcrdId(Table, , "record_no"))
    .Fields!Reference = Reference
    .Fields!war_type = Title
    .Fields!Description = Description
    .Update
End With
End Sub

Public Function ExecErr(Prompt As String, _
                        Optional PromptFld As String, _
                        Optional Table As String, _
                        Optional RcrdFld As String, _
                        Optional RcrdStr As String) As String
Dim Rcrds As String
If Table <> Empty Then
    RunSql "Select * from " & Table & " where " & RcrdFld & " = '" & RcrdStr & "'"
    With Rs
        While Not .EOF = True
            Rcrds = Rcrds & .Fields(PromptFld) & "; "
            .MoveNext
        Wend
            ExecErr = "Error: " & Prompt & vbNewLine & vbNewLine & _
                "Related Records: " & Rcrds
    End With
Else
    ExecErr = Prompt
End If
End Function

