VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} dtaGroups 
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   _ExtentX        =   11483
   _ExtentY        =   13600
   FolderFlags     =   5
   TypeLibGuid     =   "{37FC7D61-EB86-4575-BFD3-A76907AF478F}"
   TypeInfoGuid    =   "{39B2E71E-6338-46E0-989C-554860CBF96F}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "conDb"
      ConnDispId      =   1038
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\subjects\Prog2\Projects\final\POS\POS.mdb;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   6
   BeginProperty Recordset1 
      CommandName     =   "ItemSummary"
      CommDispId      =   1145
      RsDispId        =   1179
      CommandText     =   "Select Distinctrow format(short_month,'mmmm') as MonthHeader, Year from tblMonthTable"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      GroupingName    =   "ItemSummary_Grouping"
      GrandTotal      =   "GrandTotal1"
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "MonthHeader"
         Caption         =   "MonthHeader"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   510
         Scale           =   0
         Type            =   204
         Name            =   "Year"
         Caption         =   "Year"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Year"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "ItemCount"
         AggOn           =   "ItemDetails"
         AggField        =   "p_code"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "ItemCount"
         Caption         =   "ItemCount"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "ItemDetails"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"dtaGroups.dsx":0000
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "ItemSummary"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "DateRelated"
         Caption         =   "DateRelated"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "p_code"
         Caption         =   "p_code"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "net_price"
         Caption         =   "net_price"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "condition"
         Caption         =   "condition"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "quantity"
         Caption         =   "quantity"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "sold"
         Caption         =   "sold"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "year"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "MonthHeader"
         ChildField      =   "DateRelated"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Year"
         ChildField      =   "year"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "ByCategory"
      CommDispId      =   1156
      RsDispId        =   1170
      CommandText     =   "SELECT description, dYear from tblType"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   16711935
         Scale           =   0
         Type            =   204
         Name            =   "dYear"
         Caption         =   "dYear"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "TotalItems"
         AggOn           =   "Details"
         AggField        =   "p_code"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "TotalItems"
         Caption         =   "TotalItems"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "Details"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"dtaGroups.dsx":00CB
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "ByCategory"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "p_code"
         Caption         =   "p_code"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "brand_type"
         Caption         =   "brand_type"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "SellingPrice"
         Caption         =   "SellingPrice"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "location"
         Caption         =   "location"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "description"
         ChildField      =   "brand_type"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "dYear"
         ChildField      =   "dYear"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "ByLocation"
      CommDispId      =   1172
      RsDispId        =   1177
      CommandText     =   "SELECT description, dYear from tblLocation"
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   510
         Scale           =   0
         Type            =   204
         Name            =   "dYear"
         Caption         =   "dYear"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   1
      BeginProperty Aggregate1 
         Name            =   "TotalItems"
         AggOn           =   "DetailsLocation"
         AggField        =   "p_code"
         AggType         =   1
         AggFunction     =   3
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "TotalItems"
         Caption         =   "TotalItems"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "DetailsLocation"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"dtaGroups.dsx":0162
      ActiveConnectionName=   "conDb"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "ByLocation"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "description"
         Caption         =   "description"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "p_code"
         Caption         =   "p_code"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "brand_type"
         Caption         =   "brand_type"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "SellingPrice"
         Caption         =   "SellingPrice"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "location"
         Caption         =   "location"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "dYear"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "description"
         ChildField      =   "location"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "dYear"
         ChildField      =   "dYear"
         ParentType      =   0
         ChildType       =   1
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "dtaGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
