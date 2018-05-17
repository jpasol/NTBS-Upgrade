VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} deCYXINQ02 
   ClientHeight    =   9255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11400
   _ExtentX        =   20108
   _ExtentY        =   16325
   FolderFlags     =   1
   TypeLibGuid     =   "{BF755FC1-047A-11D3-9F1A-00105A626E67}"
   TypeInfoGuid    =   "{BF755FC2-047A-11D3-9F1A-00105A626E67}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Billing"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Billing;Data Source=SBITCBILLING"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   19
   BeginProperty Recordset1 
      CommandName     =   "getInformation"
      CommDispId      =   1002
      RsDispId        =   1007
      CommandText     =   "select username = user_name(),workstation = host_name()"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   128
         Scale           =   0
         Type            =   202
         Name            =   "username"
         Caption         =   "username"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   128
         Scale           =   0
         Type            =   202
         Name            =   "workstation"
         Caption         =   "workstation"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "GetUserType"
      CommDispId      =   1079
      RsDispId        =   1087
      CommandText     =   "SELECT offcde FROM userinfo WHERE userid = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "offcde"
         Caption         =   "offcde"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pUserid"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "GetDate"
      CommDispId      =   1088
      RsDispId        =   -1
      CommandText     =   "dbo.up_getsysdate"
      ActiveConnectionName=   "Billing"
      CallSyntax      =   "{? = CALL dbo.up_getsysdate( ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@@pDATE"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "GetTellerTotals"
      CommDispId      =   1090
      RsDispId        =   1098
      CommandText     =   $"deCYXINQ02.dsx":0000
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field2 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "Cash"
         Caption         =   "Cash"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "Adr"
         Caption         =   "Adr"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "Collection"
         Caption         =   "Collection"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "ArrastreAmt"
      CommDispId      =   1099
      RsDispId        =   1105
      CommandText     =   $"deCYXINQ02.dsx":0177
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "ArrastreAmtUG"
      CommDispId      =   1106
      RsDispId        =   1153
      CommandText     =   $"deCYXINQ02.dsx":021F
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "ArrastreVat"
      CommDispId      =   1111
      RsDispId        =   1116
      CommandText     =   $"deCYXINQ02.dsx":02C7
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "ArrastreVatUG"
      CommDispId      =   1117
      RsDispId        =   1154
      CommandText     =   $"deCYXINQ02.dsx":035D
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "ArrastreNV"
      CommDispId      =   1123
      RsDispId        =   1155
      CommandText     =   $"deCYXINQ02.dsx":03F3
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "ArrastreNVUG"
      CommDispId      =   1129
      RsDispId        =   1156
      CommandText     =   $"deCYXINQ02.dsx":049B
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "ArrastreWtx"
      CommDispId      =   1135
      RsDispId        =   1157
      CommandText     =   $"deCYXINQ02.dsx":0542
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "ArrastreWtxUG"
      CommDispId      =   1141
      RsDispId        =   1158
      CommandText     =   $"deCYXINQ02.dsx":05E9
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset13 
      CommandName     =   "Wharfage"
      CommDispId      =   1147
      RsDispId        =   1159
      CommandText     =   "SELECT out = SUM(whfamt) FROM ccrcyx WHERE whfcde = '0' AND status <> 'CAN' AND sysdttm >= ? AND sysdttm <= ? AND userid = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset14 
      CommandName     =   "RemitCCRPay"
      CommDispId      =   1160
      RsDispId        =   -1
      CommandText     =   "UPDATE ccrpay SET updcde = 'Y' WHERE sysdttm >= ? AND sysdttm <= ? AND userid = ? AND ccrtyp = '1' AND status <> 'CAN'"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pFromDte"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "pToDte"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "pUserid"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset15 
      CommandName     =   "RemitCCRCyx"
      CommDispId      =   1172
      RsDispId        =   -1
      CommandText     =   "UPDATE ccrcyx SET updcde = 'Y' WHERE sysdttm >= ? AND sysdttm <= ? AND userid = ? AND status <> 'CAN'"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pFromDte"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "pToDte"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "pUserid"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset16 
      CommandName     =   "Container"
      CommDispId      =   1174
      RsDispId        =   1179
      CommandText     =   "SELECT * FROM ccrcyx WHERE cntnum = ? "
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   40
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
         Precision       =   1
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "itmnum"
         Caption         =   "itmnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field6 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field10 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "whfamt"
         Caption         =   "whfamt"
      EndProperty
      BeginProperty Field11 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "arramt"
         Caption         =   "arramt"
      EndProperty
      BeginProperty Field12 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field13 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrvat"
         Caption         =   "arrvat"
      EndProperty
      BeginProperty Field15 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrtax"
         Caption         =   "arrtax"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field17 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzl"
         Caption         =   "cntovzl"
      EndProperty
      BeginProperty Field18 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzw"
         Caption         =   "cntovzw"
      EndProperty
      BeginProperty Field19 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzh"
         Caption         =   "cntovzh"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ovzums"
         Caption         =   "ovzums"
      EndProperty
      BeginProperty Field21 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "trncde"
         Caption         =   "trncde"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "whfcde"
         Caption         =   "whfcde"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field25 
         Precision       =   5
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dolrte"
         Caption         =   "dolrte"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "commod"
         Caption         =   "commod"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "trknam"
         Caption         =   "trknam"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "pltnum"
         Caption         =   "pltnum"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "trkchs"
         Caption         =   "trkchs"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field35 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ovrccr"
         Caption         =   "ovrccr"
      EndProperty
      BeginProperty Field36 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ppanum"
         Caption         =   "ppanum"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset17 
      CommandName     =   "CCR"
      CommDispId      =   1180
      RsDispId        =   1186
      CommandText     =   "SELECT * FROM ccrcyx WHERE ccrnum = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   40
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
         Precision       =   1
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "itmnum"
         Caption         =   "itmnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field6 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field10 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "whfamt"
         Caption         =   "whfamt"
      EndProperty
      BeginProperty Field11 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "arramt"
         Caption         =   "arramt"
      EndProperty
      BeginProperty Field12 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field13 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrvat"
         Caption         =   "arrvat"
      EndProperty
      BeginProperty Field15 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrtax"
         Caption         =   "arrtax"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field17 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzl"
         Caption         =   "cntovzl"
      EndProperty
      BeginProperty Field18 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzw"
         Caption         =   "cntovzw"
      EndProperty
      BeginProperty Field19 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzh"
         Caption         =   "cntovzh"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ovzums"
         Caption         =   "ovzums"
      EndProperty
      BeginProperty Field21 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "trncde"
         Caption         =   "trncde"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "whfcde"
         Caption         =   "whfcde"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field25 
         Precision       =   5
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dolrte"
         Caption         =   "dolrte"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "commod"
         Caption         =   "commod"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "trknam"
         Caption         =   "trknam"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "pltnum"
         Caption         =   "pltnum"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "trkchs"
         Caption         =   "trkchs"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field35 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ovrccr"
         Caption         =   "ovrccr"
      EndProperty
      BeginProperty Field36 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ppanum"
         Caption         =   "ppanum"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
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
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset18 
      CommandName     =   "AdrAmount"
      CommDispId      =   1187
      RsDispId        =   1192
      CommandText     =   "SELECT out = SUM(adramt) FROM ccrpay WHERE status <> 'CAN' AND sysdttm >= ? AND sysdttm <= ? AND userid = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "out"
         Caption         =   "out"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pFromDate"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "pToDate"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "pUserid"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset19 
      CommandName     =   "TotalPOS"
      CommDispId      =   1193
      RsDispId        =   1198
      CommandText     =   $"deCYXINQ02.dsx":0690
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field2 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "POS"
         Caption         =   "POS"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "deCYXINQ02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
