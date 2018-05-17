VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} deCYCCR 
   ClientHeight    =   11295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20550
   _ExtentX        =   36248
   _ExtentY        =   19923
   FolderFlags     =   7
   TypeLibGuid     =   "{0917860D-1236-11D3-BD7D-00105A64485A}"
   TypeInfoGuid    =   "{0917860E-1236-11D3-BD7D-00105A64485A}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Billing"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Password=Ictsi123;Persist Security Info=True;User ID=SA_ICTSI;Initial Catalog=billing;Data Source=SBITCBILLING"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   4
   BeginProperty Recordset1 
      CommandName     =   "getCCRList"
      CommDispId      =   1002
      RsDispId        =   1040
      CommandText     =   $"deCYCCR.dsx":0000
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cusnam"
         Caption         =   "cusnam"
      EndProperty
      BeginProperty Field2 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Reference"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "getCCRDetails"
      CommDispId      =   1010
      RsDispId        =   1041
      CommandText     =   "Select * from CCRdtl where refnum = ? and seqnum = ? and status <> 'CAN'"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   41
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "itmnum"
         Caption         =   "itmnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ccrtyp"
         Caption         =   "ccrtyp"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   129
         Name            =   "chargetyp"
         Caption         =   "chargetyp"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descr"
         Caption         =   "descr"
      EndProperty
      BeginProperty Field8 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "docrefno"
         Caption         =   "docrefno"
      EndProperty
      BeginProperty Field9 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "regnum"
         Caption         =   "regnum"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field12 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "amt"
         Caption         =   "amt"
      EndProperty
      BeginProperty Field15 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "vatamt"
         Caption         =   "vatamt"
      EndProperty
      BeginProperty Field16 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "wtax"
         Caption         =   "wtax"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "stostat"
         Caption         =   "stostat"
      EndProperty
      BeginProperty Field19 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "lngth"
         Caption         =   "lngth"
      EndProperty
      BeginProperty Field20 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "width"
         Caption         =   "width"
      EndProperty
      BeginProperty Field21 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "height"
         Caption         =   "height"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ums"
         Caption         =   "ums"
      EndProperty
      BeginProperty Field23 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "quantity"
         Caption         =   "quantity"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field25 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field26 
         Precision       =   7
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field27 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field28 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "enrfrdttm"
         Caption         =   "enrfrdttm"
      EndProperty
      BeginProperty Field29 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "enstodttm"
         Caption         =   "enstodttm"
      EndProperty
      BeginProperty Field30 
         Precision       =   4
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "stordys"
         Caption         =   "stordys"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "rectag"
         Caption         =   "rectag"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   129
         Name            =   "shplin"
         Caption         =   "shplin"
      EndProperty
      BeginProperty Field36 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "pod"
         Caption         =   "pod"
      EndProperty
      BeginProperty Field38 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field39 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field41 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "getTotal"
      CommDispId      =   1019
      RsDispId        =   1053
      CommandText     =   $"deCYCCR.dsx":00B7
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "TotalAmt"
         Caption         =   "TotalAmt"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "getAdramt"
      CommDispId      =   1028
      RsDispId        =   1033
      CommandText     =   "select * from CCRpay where refnum = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   13
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   129
         Name            =   "cuscde"
         Caption         =   "cuscde"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cusnam"
         Caption         =   "cusnam"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cshamt"
         Caption         =   "cshamt"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "adramt"
         Caption         =   "adramt"
      EndProperty
      BeginProperty Field6 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "adrnum"
         Caption         =   "adrnum"
      EndProperty
      BeginProperty Field7 
         Precision       =   9
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chgamt"
         Caption         =   "chgamt"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "rectag"
         Caption         =   "rectag"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ccrtyp"
         Caption         =   "ccrtyp"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Reference"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "deCYCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
