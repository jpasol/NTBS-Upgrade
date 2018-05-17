Attribute VB_Name = "cConnSparcs"
Public RfrPlugIn As Date

Public Function Sparcs_LastDisch(ByVal pContNum As String, ByVal pCharge As String, ByVal pPaid As String, ByVal pGKey As String)
 Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
    'strAuthorization = "Basic bjRhcGk6d2VsY29tZQ==" 'c3NhbmNoZXo6cGFzc3dvcmQ="
    'strAuthorization = "Basic c3NhbmNoZXo6cGFzc3dvcmQ="
    'strUrl = "http://172.16.0.219:9080/apex/services/inventoryservice?wsdl"
    'strSoapAction = "POST http://172.16.0.219:9080/apex/services/inventoryservice HTTP/1.1"
    
    'strUrl = "http://192.168.11.151:9080/apex/services/inventoryservice?wsdl"
    ' strSoapAction = "POST http://192.168.11.151:9080/apex/services/inventoryservice HTTP/1.1"
    
    strUrl = strN4Url & "/apex/services/inventoryservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/inventoryservice HTTP/1.1"
  
  strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
"   <soapenv:Header/>" & _
"   <soapenv:Body>" & _
"      <inv:proposePaidThruDay>" & _
"        <inv:scopeCoordinateIdsWsType>" & _
"            <!--Optional:-->" & _
"            <v1:operatorId>ICTSI</v1:operatorId>" & _
"            <!--Optional:-->" & _
"            <v1:complexId>PH</v1:complexId>" & _
"            <!--Optional:-->" & _
"            <v1:facilityId>SBITC</v1:facilityId>" & _
"            <!--Optional:-->" & _
"            <v1:yardId>SBITC</v1:yardId>" & _
"         </inv:scopeCoordinateIdsWsType>" & _
"         <inv:eqId>"

strChargeFor = "</inv:eqId>" & _
"         <inv:chargeFor>"

strPaid = "</inv:chargeFor>" & _
"         <inv:paidThruDay>"

strGKey = "</inv:paidThruDay>" & _
"         <inv:extractGkey>"

strSoapEnd = "</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd


    ' Call PostWebservice and put result in text box
    If pCharge = "STORAGE" Then
        strOutput = GetDischarge(strUrl, strSoapAction, strXML, strN4Authorization)
    ElseIf pCharge = "REEFER" Then
        strOutput = GetReefer(strUrl, strSoapAction, strXML, strN4Authorization)
    End If
End Function

Public Function GetDischarge(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim strQueryDays As String
    Dim result As String
    Dim resultDays As String
    Dim bError As Boolean
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    'strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns5:previousPaidThruDay"
'    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns3:lastFreeDay"
'    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

'Added
'    strQueryDays = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns9:daysPaid"
'    resultDays = objResult.selectSingleNode(strQueryDays).Text
'    If CInt(resultDays) > 0 Then
'        strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns5:previousPaidThruDay"
'        result = Left(objResult.selectSingleNode(strQuery).Text, 10)
'    ElseIf CInt(resultDays) = 0 Then
'        strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns3:lastFreeDay"
'        result = Left(objResult.selectSingleNode(strQuery).Text, 10)
'    End If

    'Edited by Navis Project Team 11/05/2009
'    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns2:firstFreeDay"
    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns5:previousPaidThruDay"
    '    bError = IsError(objResult.selectSingleNode(strQuery).Text)
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)
    
    dStorage = result

    Exit Function
    
Err_PW:
    

'On Error GoTo Err_PW2

'        strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns3:lastFreeDay"
'        result = Left(objResult.selectSingleNode(strQuery).Text, 10)
'        dStorage = result
'        Exit Function
        Call GetDischargeLastFreeDay(objResult)
'Err_PW2:
'    GetDischarge = "Error: " & Err.Number & " - " & Err.Description
'    result = "1899-12-30 00:00:00"
'    dStorage = result
'
'    ' Close object
    Set objXmlHttp = Nothing
'
'    ' Return result
'    GetDischarge = strRet

End Function

Public Function GetDischargeLastFreeDay(ByVal objResult As Object) As String
Dim result As String
On Error GoTo ErrHandler
    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns3:lastFreeDay"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)
    dStorage = result
    GetDischargeLastFreeDay = result
    Set objXmlHttp = Nothing
    Exit Function
ErrHandler:
result = "1899-12-30 00:00:00"
dStorage = result
Set objXmlHttp = Nothing
GetDischargeLastFreeDay = "Error: " & Err.Number & " - " & Err.Description
End Function
Public Function GetReefer(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String

    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    
    'added Navis Project Team 10/28/2009
'    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns4:proposedPaidThruDay"
'    result = Left(objResult.selectSingleNode(strQuery).Text, 10) & " " & Mid(objResult.selectSingleNode(strQuery).Text, 12, 8)
'
'    dReefer = result

    Dim nPaidThruPos As Integer
    Dim bHasPaidThru As Boolean
    Dim bHasPlugIn As Boolean
    
    bHasPaidThru = False
    bHasPlugIn = False
    'Edited by Navis Project Team 11/05/2009
'    nPaidThruPos = InStr(XmlBody, "PowerPaidThruTime:")
'    If nPaidThruPos > 0 Then
'        strQueryPaidThru = Mid(XmlBody, nPaidThruPos)
'        If IsDate(Mid(strQueryPaidThru, 19, 19)) Then
'            result = Mid(strQueryPaidThru, 19, 19) 'objResult.selectSingleNode(strQuery).Text
'            dReefer = result
'            bHasPaidThru = True
'        End If
'    End If
'   If bHasPaidThru = False Then
'        strQuery = Mid(XmlBody, 701, 37)
'        If Left(strQuery, 17) = "PowerConnectTime:" Then
'            If IsDate(Mid(strQuery, 18, 19)) Then
'                result = Mid(strQuery, 18, 19) 'objResult.selectSingleNode(strQuery).Text
'                dReefer = result
'                bHasPlugIn = True
'            End If
'        End If
'    End If
    nPaidThruPos = InStr(strRet, "PowerPaidThruTime:")
    If nPaidThruPos > 0 Then
        strQueryPaidThru = Mid(strRet, nPaidThruPos)
        If IsDate(Mid(strQueryPaidThru, 19, 19)) Then
            result = Mid(strQueryPaidThru, 19, 19) 'objResult.selectSingleNode(strQuery).Text
            dReefer = result
            bHasPaidThru = True
        End If
    End If
   If bHasPaidThru = False Then
        strQuery = Mid(strRet, 701, 37)
        If Left(strQuery, 17) = "PowerConnectTime:" Then
            If IsDate(Mid(strQuery, 18, 19)) Then
                result = Mid(strQuery, 18, 19) 'objResult.selectSingleNode(strQuery).Text
                dReefer = result
                bHasPlugIn = True
            End If
        End If
    End If
    If bHasPaidThru = False And bHasPlugIn = False Then
        dReefer = "1899-12-30 00:00:00"
    End If

    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetReefer = strRet
    Exit Function
    
Err_PW:
MsgBox Err.Description
    GetReefer = "Error: " & Err.Number & " - " & Err.Description

End Function

Public Function Sparcs_DGCode(ByVal strContNum As String, ByVal strParamCat As String)
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam As String  ', strParamCat
    Dim strOutput As String
    
    strOutput = ""
'    strAuthorization = "Basic YWRtaW46cGFzc3dvcmQ="
    'strFilter = "http://172.16.0.219:9080/apex/api/query?filtername=GETDGCODE&PARM_ContNum="
    'strFilter = "http://192.168.11.151:9080/apex/api/query?filtername=GETDGCODE&PARM_ContNum="
    strFilter = strN4Url & "/apex/api/query?filtername=GETDGCODE&PARM_ContNum="
    'strContNum = strParam
    strCategory = "&PARM_Category="
    'strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    'strSoapAction = "POST http://172.16.0.219:9080/apex/services/inventoryservice HTTP/1.1"
    'strSoapAction = "POST http://192.168.11.152:9080/apex/services/inventoryservice HTTP/1.1"
    strSoapAction = "POST " & strN4Url & "/apex/services/inventoryservice HTTP/1.1"
  

'    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
'"   <soapenv:Header/>" & _
'"   <soapenv:Body>" & _
'"      <inv:proposePaidThruDay>" & _
'"        <inv:scopeCoordinateIdsWsType>" & _
'"            <!--Optional:-->" & _
'"            <v1:operatorId>ICTSI</v1:operatorId>" & _
'"            <!--Optional:-->" & _
'"            <v1:complexId>PH</v1:complexId>" & _
'"            <!--Optional:-->" & _
'"            <v1:facilityId>SBITC</v1:facilityId>" & _
'"            <!--Optional:-->" & _
'"            <v1:yardId>SBITC</v1:yardId>" & _
'"         </inv:scopeCoordinateIdsWsType>" & _
'"         <inv:eqId>APHU6623529</inv:eqId>" & _
'"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
'"         <inv:paidThruDay>?</inv:paidThruDay>" & _
'"         <inv:extractGkey>4992</inv:extractGkey>" & _
'"      </inv:proposePaidThruDay>" & _
'"   </soapenv:Body>" & _
'"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
'    strOutput = GetDGCode(strUrl, strSoapAction, strXML, strN4Authorization)
    strOutput = GetDGCode(strUrl, strSoapAction, "", strN4Authorization)
'    ReleaseDG (strContNum)
End Function

Public Function GetDGCode(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String

    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "GET", AsmxUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext
    'MsgBox (strRet)

    objResult.async = False
    objResult.loadxml strRet
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
    'Text1.Text = result
    
    ParseDG (result)
    
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetDGCode = strRet
    Exit Function
    
Err_PW:
    GetDGCode = "Error: " & Err.Number & " - " & Err.Description

End Function

Public Function Sparcs_OOG(ByVal strContNum As String, ByVal strParamCat As String)
     
    Dim strSoapAction As String
    Dim strUrl, strFilter, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    
    'strParam = "SHAR1809099"
    strOutput = ""
'    strAuthorization = "Basic YWRtaW46cGFzc3dvcmQ="
    'strUrl = "http://172.16.0.219:9080/apex/api/query?filtername=GETOOG&PARM_ContNum=SHAR1809099&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    'strFilter = "http://172.16.0.219:9080/apex/api/query?filtername=GETOOG&PARM_ContNum="
    'strFilter = "http://192.168.11.151:9080/apex/api/query?filtername=GETOOG&PARM_ContNum="
    strFilter = strN4Url & "/apex/api/query?filtername=GETOOG&PARM_ContNum="
    'strContNum = strParam
    strCategory = "&PARM_Category="
    'strParamCat = "IMPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    'strSoapAction = "POST http://172.16.0.219:9080/apex/services/inventoryservice HTTP/1.1"
     'strSoapAction = "POST http://192.168.11.151:9080/apex/services/inventoryservice HTTP/1.1"
     strSoapAction = "POST " & strN4Url & "/apex/services/inventoryservice HTTP/1.1"
  

'    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
'"   <soapenv:Header/>" & _
'"   <soapenv:Body>" & _
'"      <inv:proposePaidThruDay>" & _
'"        <inv:scopeCoordinateIdsWsType>" & _
'"            <!--Optional:-->" & _
'"            <v1:operatorId>ICTSI</v1:operatorId>" & _
'"            <!--Optional:-->" & _
'"            <v1:complexId>PH</v1:complexId>" & _
'"            <!--Optional:-->" & _
'"            <v1:facilityId>SBITC</v1:facilityId>" & _
'"            <!--Optional:-->" & _
'"            <v1:yardId>SBITC</v1:yardId>" & _
'"         </inv:scopeCoordinateIdsWsType>" & _
'"         <inv:eqId>APHU6623529</inv:eqId>" & _
'"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
'"         <inv:paidThruDay>?</inv:paidThruDay>" & _
'"         <inv:extractGkey>4992</inv:extractGkey>" & _
'"      </inv:proposePaidThruDay>" & _
'"   </soapenv:Body>" & _
'"</soapenv:Envelope>"

'    strOutput = GetOOG(strUrl, strSoapAction, strXML, strN4Authorization)
    strOutput = GetOOG(strUrl, strSoapAction, "", strN4Authorization)
'    ReleaseOOG (strContNum)
End Function

Public Function GetOOG(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String

    On Error GoTo Err_PW
    
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    objDom.async = False
    objDom.loadxml XmlBody

    objXmlHttp.Open "GET", AsmxUrl, False, strN4UserName, strN4Password
    
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    objXmlHttp.send objDom.xml

    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
    'Text2.Text = result
    
    ParseOOG (result)
    
    'frmManifestCont.mskOVLength.Text = length
    'frmManifestCont.mskOVWidth.Text = width
    'frmManifestCont.mskOVHeight.Text = height
    
    Set objXmlHttp = Nothing
    
 
    GetOOG = strRet
    
    Exit Function
    
Err_PW:
    GetOOG = "Error: " & Err.Number & " - " & Err.Description
End Function

Public Function ParseOOG(ByVal sName As String)
Dim sLetter As String
Dim iAsc As Integer
Dim i As Integer
Dim ctrInt, ctrNo As Integer
Dim iNum1, iNum2 As Integer

ctrInt = 1
ctrNo = 1
If sName <> "" Then
    For i = 1 To Len(sName)
        sLetter = Mid(sName, 1, 1)
        iAsc = Asc(sLetter)
        
        If (iAsc >= 48 And iAsc <= 57) Then
            If ctrNo = 1 Then
                If ctrInt = 1 Then
                    iNum1 = iNum1 & Left(sName, 1) 'sLetter
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                ElseIf ctrInt = 2 Then
                    iNum2 = iNum2 & Left(sName, 1)
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                End If
            ElseIf ctrNo = 2 Then
                If ctrInt = 1 Then
                    iNum1 = iNum1 & Left(sName, 1) 'sLetter
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                ElseIf ctrInt = 2 Then
                    iNum2 = iNum2 & Left(sName, 1)
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                End If
            ElseIf ctrNo = 3 Then
                If ctrInt = 1 Then
                    iNum1 = iNum1 & Left(sName, 1) 'sLetter
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                ElseIf ctrInt = 2 Then
                    iNum2 = iNum2 & Left(sName, 1)
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                End If
            End If
            
        Else
            If iAsc = 43 Then
                sName = Replace(sName, Left(sName, 1), "", 1, 1)
                ctrInt = ctrInt + 1
            Else
                If ctrNo = 1 Then
                    If frmCYSCCR.tabTran.Tab = 0 Then
                        frmCYSCCR.txtARROvzLen.Text = iNum1 + iNum2
                    ElseIf frmCYSCCR.tabTran.Tab = 1 Then
                        frmCYSCCR.txtStoOvzLen.Text = iNum1 + iNum2
                    ElseIf frmCYSCCR.tabTran.Tab = 5 Then
                        frmCYSCCR.txtOthOvzLen.Text = iNum1 + iNum2
                    End If
                    iNum1 = Empty
                    iNum2 = Empty
                    ctrNo = ctrNo + 1
                    ctrInt = 1
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                ElseIf ctrNo = 2 Then
                    If frmCYSCCR.tabTran.Tab = 0 Then
                        frmCYSCCR.txtARROvzWid.Text = iNum1 + iNum2
                    ElseIf frmCYSCCR.tabTran.Tab = 1 Then
                        frmCYSCCR.txtStoOvzWid.Text = iNum1 + iNum2
                    ElseIf frmCYSCCR.tabTran.Tab = 5 Then
                        frmCYSCCR.txtOthOvzWid.Text = iNum1 + iNum2
                    End If
    '                frmCYSCCR.txtARROvzWid.Text = iNum1 + iNum2
                    iNum1 = Empty
                    iNum2 = Empty
                    ctrNo = ctrNo + 1
                    ctrInt = 1
                    sName = Replace(sName, Left(sName, 1), "", 1, 1)
                End If
            End If
        End If
    
    Next
    'height
    If frmCYSCCR.tabTran.Tab = 0 Then
        frmCYSCCR.txtARROvzHgt.Text = iNum1 + iNum2
    ElseIf frmCYSCCR.tabTran.Tab = 1 Then
        frmCYSCCR.txtStoOvzHgt.Text = iNum1 + iNum2
    ElseIf frmCYSCCR.tabTran.Tab = 5 Then
        frmCYSCCR.txtOthOvzHgt.Text = iNum1 + iNum2
    End If
    'frmCYSCCR.txtARROvzHgt.Text = iNum1 + iNum2
    iNum1 = Empty
    iNum2 = Empty
    ctrNo = ctrNo + 1
    ctrInt = 1
    sName = Replace(sName, Left(sName, 1), "", 1, 1)

End If
End Function

'Added by Navis Project Team 10/28/2009
Public Function ParseDG(ByVal sName As String)
Dim strDG As String
Dim arrDGList() As String
Dim arrDGVal() As String
Dim iCtrDG As Integer
Dim iCtrDGVal As Integer
Dim intHighestDG As Integer
Dim comboCounter As Integer
'sharon 05Nov2009 begin
Dim arrRate1(10) As String
Dim arrRate2(10) As String
Dim arrRate3(10) As String
Dim iR1Ctr, iR2Ctr, iR3Ctr As Integer
iR1Ctr = 0
iR2Ctr = 0
iR3Ctr = 0
'sharon 05Nov2009 end

intHighestDG = 0

arrDGList = Split(sName, ",")
If UBound(arrDGList) >= 0 Then
    For iCtrDG = 0 To UBound(arrDGList)
'sharon 05Nov2009        If intHighestDG < CDbl(arrDGList(iCtrDG)) Then
'sharon 05Nov2009            arrDGVal = Split(arrDGList(iCtrDG), ".")
'sharon 05Nov2009            intHighestDG = CInt(arrDGVal(0))
'sharon 05Nov2009        End If
            arrDGVal = Split(arrDGList(iCtrDG), ".")
            intHighestDG = CInt(arrDGVal(0))
            If intHighestDG = 1 Or intHighestDG = 6 Or intHighestDG = 8 Then
                arrRate1(iR1Ctr) = intHighestDG
                iR1Ctr = iR1Ctr + 1
                iCtrDG = UBound(arrDGList)
            ElseIf intHighestDG = 2 Or intHighestDG = 3 Or intHighestDG = 4 Or intHighestDG = 7 Then
                arrRate2(iR2Ctr) = intHighestDG
                iR2Ctr = iR2Ctr + 1
            ElseIf intHighestDG = 5 Or intHighestDG = 9 Then
                arrRate3(iR3Ctr) = intHighestDG
                iR3Ctr = iR3Ctr + 1
            End If
    Next
    If iR1Ctr > 0 Then
        intHighestDG = CInt(arrRate1(0))
    ElseIf iR2Ctr > 0 Then
        intHighestDG = CInt(arrRate2(0))
    ElseIf iR3Ctr > 0 Then
        intHighestDG = CInt(arrRate3(0))
    End If
End If

frmCYSCCR.cboDanger.ListIndex = intHighestDG
End Function

'Modified by Navis Project Team 10/28/2009
Public Function GetGKey(ByVal pCont As String, ByVal pStatus As String, ByVal pType As String, Optional ByVal pCat As String = "", Optional ByVal pVisit As String = "") As String
    Dim rstGKey As ADODB.Recordset
    Dim strGKey As String
    Dim strResult As String
    Set rstGKey = New ADODB.Recordset
    
    strResult = ""
    strGKey = ""
    If pVisit <> "" And pCat <> "" Then
        strGKey = "SELECT top 1 gkey, event_start_time FROM argo_chargeable_unit_events " & _
                    "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' " & _
                    "and event_type = '" & pType & "' " & _
                    "category = '" & pCat & "' and ib_id= '" & pVisit & "'" & _
                    "order by changed desc"
    ElseIf pVisit = "" And pCat = "" Then
        strGKey = "SELECT top 1 gkey, event_start_time FROM argo_chargeable_unit_events " & _
                    "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' " & _
                    "and event_type = '" & pType & "' " & _
                    "order by changed desc"
    End If



    rstGKey.Open strGKey, gcnnNavis, adOpenForwardOnly, adLockReadOnly

    If Not rstGKey.BOF = True Or Not rstGKey.EOF = True Then
        strResult = rstGKey.Fields(0)
        RfrPlugIn = rstGKey.Fields(1)
    End If
    GetGKey = strResult
End Function

'Added by Navis Project Team 11/05/2009
'Direct database query for retrieving PaidThruDay for reefer and storage
Public Function GetLastDischargeDate(ByVal pCont As String, ByVal pStatus As String, ByVal pType As String, Optional ByVal pCat As String = "", Optional ByVal pVisit As String = "") As String
    Dim rstDischargeData As ADODB.Recordset
    Dim strQuery As String
    Dim strResult As String
    Set rstDischargeData = New ADODB.Recordset
    
    strResult = ""
    strQuery = ""
    If pVisit <> "" And pCat <> "" Then
        strQuery = "SELECT top 1 event_start_time, paid_thru_day FROM argo_chargeable_unit_events " & _
                    "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' " & _
                    "and event_type = '" & pType & "' " & _
                    "category = '" & pCat & "' and ib_id= '" & pVisit & "'" & _
                    "order by changed desc"
    ElseIf pVisit = "" And pCat = "" Then
        strQuery = "SELECT top 1 event_start_time, paid_thru_day FROM argo_chargeable_unit_events " & _
                    "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' " & _
                    "and event_type = '" & pType & "' " & _
                    "order by changed desc"
    End If

    rstDischargeData.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly

    If Not rstDischargeData.BOF = True Or Not rstDischargeData.EOF = True Then
        If pType = "REEFER" Then
            If rstDischargeData.Fields(1) <> "" Then
                strResult = rstDischargeData.Fields(1)
            Else
                strResult = rstDischargeData.Fields(0)
            End If
        ElseIf pType = "STORAGE" Then
            strResult = rstDischargeData.Fields(1)
        End If
    End If
    GetLastDischargeDate = strResult
    Set rstDischargeData = Nothing
End Function

'Added by Navis Project Team 11/05/2009
'Direct database query for retrieving PaidThruDay for reefer and storage
Public Sub GetReeferDates(ByVal pCont As String, ByVal pStatus As String, ByRef PluginDate As String, ByRef PaidThruDate As String, Optional ByVal pCat As String = "", Optional ByVal pVisit As String = "")
    Dim rstDischargeData As ADODB.Recordset
    Dim strQuery As String
    Dim strResult As String
    Dim Unit_GKey As String
    Set rstDischargeData = New ADODB.Recordset
    
    strResult = ""
    strQuery = ""
    If pVisit <> "" And pCat <> "" Then
        strQuery = "SELECT top 1 event_start_time, paid_thru_day, unit_gkey FROM argo_chargeable_unit_events " & _
                    "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' " & _
                    "and event_type = 'REEFER' " & _
                    "category = '" & pCat & "' and ib_id= '" & pVisit & "'" & _
                    "order by changed desc"
    ElseIf pVisit = "" And pCat = "" Then
        strQuery = "SELECT top 1 event_start_time, paid_thru_day, unit_gkey FROM argo_chargeable_unit_events " & _
                    "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' " & _
                    "and event_type = 'REEFER' " & _
                    "order by changed desc"
    End If

    rstDischargeData.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly

    If Not rstDischargeData.BOF = True Or Not rstDischargeData.EOF = True Then
        Unit_GKey = rstDischargeData.Fields(2)
        If GKeyHasUnitOut(Unit_GKey) = False Then
            PluginDate = rstDischargeData.Fields(0)
            PaidThruDate = rstDischargeData.Fields(1)
        End If
    End If
    Set rstDischargeData = Nothing
End Sub

Public Function GKeyHasUnitOut(Unit_GKey As String) As Boolean
    Dim rsGKey As ADODB.Recordset
    Dim strQuery As String
    Dim bResult As Boolean

    Set rsGKey = New ADODB.Recordset

    bResult = False
    
    strQuery = "SELECT unit_gkey FROM argo_chargeable_unit_events WHERE unit_gkey=" & Unit_GKey & " AND event_type LIKE 'UNIT_OUT%' AND status <> 'CANCELLED' "
    
    rsGKey.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    If Not rsGKey.BOF Or Not rsGKey.EOF Then
        bResult = True
    End If
    
    rsGKey.Close
    Set rsGKey = Nothing
    
    GKeyHasUnitOut = bResult
End Function


Public Sub GetContainerLastestCategory(ByVal ContainerNo As String, ByRef Category As String, ByRef HasUnitOut As Boolean)
    Dim rsGKey As ADODB.Recordset
    Dim strQuery As String
    Dim Unit_GKey As String
    
    Set rsGKey = New ADODB.Recordset

    bResult = False
    Category = ""
    strQuery = "SELECT TOP 1 category, unit_gkey FROM argo_chargeable_unit_events WHERE equipment_id='" & ContainerNo & "' AND status <> 'CANCELLED' ORDER BY created DESC "
    
    rsGKey.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    If Not rsGKey.BOF Or Not rsGKey.EOF Then
        Category = rsGKey.Fields(0)
        Unit_GKey = rsGKey.Fields(1)
    End If
    
    HasUnitOut = GKeyHasUnitOut(Unit_GKey)
    
    rsGKey.Close
    Set rsGKey = Nothing
End Sub

'public GetUnitCreateDate
'Public Sub GetUnitCreateDate(ByVal pCont As String, Optional ByVal pCat As String = "", Optional ByVal pVisit As String = "")
'    Dim rstDischargeData As ADODB.Recordset
'    Dim strQuery As String
'    Dim strResult As String
'    Dim strSelect As String
'    Dim strWhere As String
'
'    Set rstDischargeData = New ADODB.Recordset
'
'    strResult = ""
'    strQuery = ""
'    strSelect = "SELECT top 1 event_start_time, paid_thru_day FROM argo_chargeable_unit_events "
'    strWhere = "Where equipment_id = '" & pCont & "' and status <> 'CANCELLED' " & _
'                "and event_type = 'UNIT_CREATE' "
'
'    If pCat <> "" Then
'        strWhere = strWhere & "category = '" & pCat & "' "
'    End If
'
'    If pVisit <> "" And pCat <> "" Then
'        strQuery = "SELECT top 1 event_start_time, paid_thru_day FROM argo_chargeable_unit_events " & _
'                    "Where equipment_id = '" & pCont & "' and status <> 'CANCELLED' " & _
'                    "and event_type = 'UNIT_CREATE' " & _
'                    "category = '" & pCat & "' and ib_id= '" & pVisit & "'" & _
'                    "order by changed desc"
'    ElseIf pVisit = "" And pCat = "" Then
'        strQuery = "SELECT top 1 event_start_time, paid_thru_day FROM argo_chargeable_unit_events " & _
'                    "Where equipment_id = '" & pCont & "' and status = 'CANCELLED' " & _
'                    "and event_type = 'REEFER' " & _
'                    "order by changed desc"
'    End If
'
'    rstDischargeData.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly
'
'    If Not rstDischargeData.BOF = True Or Not rstDischargeData.EOF = True Then
'        PluginDate = rstDischargeData.Fields(0)
'        PaidThruDate = rstDischargeData.Fields(1)
'    End If
'    Set rstDischargeData = Nothing
'End Sub

'this function is candidate for removal
Public Function SaveStorage(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String

    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    strQuery = "//soapenv:Envelope/soapenv:Body/updatePaidThruDayResponse/updatePaidThruDayResponse/ns1:StatusDescription"
'    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    'dStorage = result
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    SaveStorage = strRet
    Debug.Print strRet
    Exit Function
    
Err_PW:
    SaveStorage = "Error: " & Err.Number & " - " & Err.Description

End Function

'this function is candidate for removal
Public Function SaveReefer(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String

    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    strQuery = "//soapenv:Envelope/soapenv:Body/updatePaidThruDayResponse/updatePaidThruDayResponse/ns1:StatusDescription"
    'result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    dStorage = result
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    SaveReefer = strRet
    Debug.Print strRet
    Exit Function
    
Err_PW:
    SaveReefer = "Error: " & Err.Number & " - " & Err.Description

End Function

Public Function ReleaseDG(ByVal pContNum As String)
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strScope As String
    Dim strUnit As String
    Dim strRet As String
    Dim result As String

On Error GoTo ERR_Handler
    strUrl = strN4Url & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/argoservice HTTP/1.1"
  
    strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:arg=""http://www.navis.com/services/argoservice"" xmlns:v1=""http://types.webservice.argo.navis.com/v1.0""> " & _
"   <soapenv:Header/> " & _
"   <soapenv:Body> " & _
"      <arg:genericInvoke> " & _
"         <arg:scopeCoordinateIdsWsType> " & _
"            <!--Optional:--> " & _
"            <v1:operatorId>ICTSI</v1:operatorId> " & _
"            <!--Optional:--> " & _
"            <v1:complexId>PH</v1:complexId> " & _
"            <!--Optional:--> " & _
"            <v1:facilityId>SBITC</v1:facilityId> " & _
"            <!--Optional:--> " & _
"            <v1:yardId>SBITC</v1:yardId> " & _
"         </arg:scopeCoordinateIdsWsType> " & _
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id=""" & pContNum

    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING_DG"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & strUnit
    Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    Set objResult = CreateObject("MSXML2.DOMDocument")

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strN4Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strRet = objXmlHttp.responsetext
    objResult.async = False
    objResult.loadxml strRet
    
    strQuery = "//soapenv:Envelope/soapenv:Body/genericInvokeResponse/genericInvokeResponse/ns1:commonResponse/ns1:Status"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    ' Close object
    Set objXmlHttp = Nothing
    ReleaseDG = result
    Exit Function
ERR_Handler:
    ReleaseDG = "Error: " & Err.Number & " - " & Err.Description
End Function

Public Function WeighHold(ByVal pContNum As String)
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor, strUnit As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
'    strAuthorization = "Basic bjRhcGk6d2VsY29tZQ=="
    strUrl = strN4Url & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/argoservice HTTP/1.1"
  
    strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:arg=""http://www.navis.com/services/argoservice"" xmlns:v1=""http://types.webservice.argo.navis.com/v1.0""> " & _
"   <soapenv:Header/> " & _
"   <soapenv:Body> " & _
"      <arg:genericInvoke> " & _
"         <arg:scopeCoordinateIdsWsType> " & _
"            <!--Optional:--> " & _
"            <v1:operatorId>ICTSI</v1:operatorId> " & _
"            <!--Optional:--> " & _
"            <v1:complexId>PH</v1:complexId> " & _
"            <!--Optional:--> " & _
"            <v1:facilityId>SBITC</v1:facilityId> " & _
"            <!--Optional:--> " & _
"            <v1:yardId>SBITC</v1:yardId> " & _
"         </arg:scopeCoordinateIdsWsType> " & _
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id=""" & pContNum

    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""WEIGH_IMP_COMPLETE"" action=""ADD_HOLD""/></flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & strUnit
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strN4Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
    'Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    'strOutput =
End Function

'Public Function ReleaseOOG(ByVal pContNum As String)
'    Dim objDom As Object
'    Dim objXmlHttp As Object
'    Dim strSoapAction As String
'    Dim strUrl As String
'    Dim strXML As String
'    Dim strParam As String
'    Dim strOutput As String
'    Dim strScope As String
'    Dim strChargeFor, strUnit As String
'    Dim strPaid As String
'    Dim strGKey As String
'    Dim strSoapEnd As String
'    strOutput = ""
''    strAuthorization = "Basic bjRhcGk6d2VsY29tZQ=="
'    strUrl = strN4Url & "/apex/services/argoservice?wsdl"
'    strSoapAction = "POST " & strN4Url & "/apex/services/argoservice HTTP/1.1"
'
'    strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:arg=""http://www.navis.com/services/argoservice"" xmlns:v1=""http://types.webservice.argo.navis.com/v1.0""> " & _
'"   <soapenv:Header/> " & _
'"   <soapenv:Body> " & _
'"      <arg:genericInvoke> " & _
'"         <arg:scopeCoordinateIdsWsType> " & _
'"            <!--Optional:--> " & _
'"            <v1:operatorId>ICTSI</v1:operatorId> " & _
'"            <!--Optional:--> " & _
'"            <v1:complexId>PH</v1:complexId> " & _
'"            <!--Optional:--> " & _
'"            <v1:facilityId>SBITC</v1:facilityId> " & _
'"            <!--Optional:--> " & _
'"            <v1:yardId>SBITC</v1:yardId> " & _
'"         </arg:scopeCoordinateIdsWsType> " & _
'"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id=""" & pContNum
'
'    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING_OOG"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
'"      </arg:genericInvoke> " & _
'"   </soapenv:Body> " & _
'"</soapenv:Envelope>"
'
'    strXML = strScope & strUnit
'Debug.Print strXML
'
'    Set objDom = CreateObject("MSXML2.DOMDocument")
'    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
'
'
'    ' Load XML
'    objDom.async = False
'    objDom.loadxml strXML
'
'    ' Open the webservice
'    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password
'
'    ' Create headings
'    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
'    objXmlHttp.setRequestHeader "Authorization", strN4Authorization
'
'    ' Send XML command
'    objXmlHttp.send objDom.xml
'    strParam = objXmlHttp.responsetext
'   Debug.Print strParam
'    ' Close object
'    Set objXmlHttp = Nothing
'
'End Function
Public Function ReleaseOOG(ByVal pContNum As String)
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strScope As String
    Dim strUnit As String
    Dim strRet As String
    Dim result As String

On Error GoTo ERR_Handler
    strUrl = strN4Url & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/argoservice HTTP/1.1"
  
    strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:arg=""http://www.navis.com/services/argoservice"" xmlns:v1=""http://types.webservice.argo.navis.com/v1.0""> " & _
"   <soapenv:Header/> " & _
"   <soapenv:Body> " & _
"      <arg:genericInvoke> " & _
"         <arg:scopeCoordinateIdsWsType> " & _
"            <!--Optional:--> " & _
"            <v1:operatorId>ICTSI</v1:operatorId> " & _
"            <!--Optional:--> " & _
"            <v1:complexId>PH</v1:complexId> " & _
"            <!--Optional:--> " & _
"            <v1:facilityId>SBITC</v1:facilityId> " & _
"            <!--Optional:--> " & _
"            <v1:yardId>SBITC</v1:yardId> " & _
"         </arg:scopeCoordinateIdsWsType> " & _
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id=""" & pContNum

    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING_OOG"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & strUnit
    Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    Set objResult = CreateObject("MSXML2.DOMDocument")

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strN4Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strRet = objXmlHttp.responsetext
    objResult.async = False
    objResult.loadxml strRet
    
    strQuery = "//soapenv:Envelope/soapenv:Body/genericInvokeResponse/genericInvokeResponse/ns1:commonResponse/ns1:Status"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    ' Close object
    Set objXmlHttp = Nothing
    ReleaseOOG = result
    Exit Function
ERR_Handler:
    ReleaseOOG = "Error: " & Err.Number & " - " & Err.Description
End Function

Public Function ReleaseBilling(ByVal pContNum As String)
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor, strUnit As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
'    strAuthorization = "Basic bjRhcGk6d2VsY29tZQ=="
    strUrl = strN4Url & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/argoservice HTTP/1.1"
  
    strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:arg=""http://www.navis.com/services/argoservice"" xmlns:v1=""http://types.webservice.argo.navis.com/v1.0""> " & _
"   <soapenv:Header/> " & _
"   <soapenv:Body> " & _
"      <arg:genericInvoke> " & _
"         <arg:scopeCoordinateIdsWsType> " & _
"            <!--Optional:--> " & _
"            <v1:operatorId>ICTSI</v1:operatorId> " & _
"            <!--Optional:--> " & _
"            <v1:complexId>PH</v1:complexId> " & _
"            <!--Optional:--> " & _
"            <v1:facilityId>SBITC</v1:facilityId> " & _
"            <!--Optional:--> " & _
"            <v1:yardId>SBITC</v1:yardId> " & _
"         </arg:scopeCoordinateIdsWsType> " & _
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id=""" & pContNum

    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & strUnit
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strN4Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function
Public Function SavePaymentToSparcs(ByVal pContNum As String, ByVal pCharge As String, ByVal pPaid As String, ByVal pGKey As String)
 Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
'    strAuthorization = "Basic bjRhcGk6d2VsY29tZQ==" 'c3NhbmNoZXo6cGFzc3dvcmQ="
    strUrl = strN4Url & "/apex/services/inventoryservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/inventoryservice HTTP/1.1"
  
  strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
"   <soapenv:Header/>" & _
"   <soapenv:Body>" & _
"      <inv:updatePaidThruDay>" & _
"        <inv:scopeCoordinateIdsWsType>" & _
"            <!--Optional:-->" & _
"            <v1:operatorId>ICTSI</v1:operatorId>" & _
"            <!--Optional:-->" & _
"            <v1:complexId>PH</v1:complexId>" & _
"            <!--Optional:-->" & _
"            <v1:facilityId>SBITC</v1:facilityId>" & _
"            <!--Optional:-->" & _
"            <v1:yardId>SBITC</v1:yardId>" & _
"         </inv:scopeCoordinateIdsWsType>" & _
"         <inv:eqId>"

strChargeFor = "</inv:eqId>" & _
"         <inv:chargeFor>"

strPaid = "</inv:chargeFor>" & _
"         <inv:paidThruDay>"

strGKey = "</inv:paidThruDay>" & _
"         <inv:extractGkey>"

strSoapEnd = "</inv:extractGkey>" & _
"      </inv:updatePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

If pCharge = "STORAGE" Then
    pPaid = Format(pPaid, "yyyy-mm-dd") & "T" & "00:00:00 +0800"
ElseIf pCharge = "REEFER" Then
    pPaid = Format(pPaid, "yyyy-mm-dd") & "T" & Format(pPaid, "hh:mm:ss") & " +0800"
End If

strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd

Debug.Print strXML
    ' Call PostWebservice and put result in text box
'Added by Navis Project Team 10/28/2009
'    If pCharge = "STORAGE" Then
'        strOutput = SaveStorage(strUrl, strSoapAction, strXML, strN4Authorization)
'    ElseIf pCharge = "REEFER" Then
'        strOutput = SaveReefer(strUrl, strSoapAction, strXML, strN4Authorization)
'    End If
    strOutput = SavePaidThruDay(strUrl, strSoapAction, strXML, strN4Authorization)
End Function
'Added by Navis Project Team 10/28/2009
Public Function SavePaidThruDay(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String

    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password '"n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext
    Debug.Print strRet
    objResult.async = False
    objResult.loadxml strRet
    strQuery = "//soapenv:Envelope/soapenv:Body/updatePaidThruDayResponse/updatePaidThruDayResponse/ns1:StatusDescription"
    'result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    'dStorage = result
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    SavePaidThruDay = strRet
    Debug.Print strRet
    Exit Function
    
Err_PW:
    SavePaidThruDay = "Error: " & Err.Number & " - " & Err.Description
MsgBox Err.Description
End Function


'sharon begin
'06Nov2009
Public Function Sparcs_GetSOC(ByVal pCont As String, ByRef pSize As String, ByRef pStorageStart As String, ByRef pStorageEnd As String, ByRef bHasPaidThruDate As Boolean) As String
    Dim strAsmxUrl, strSoapActionUrl, strAuthorization, strXmlBody As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim strQuery As String
    Dim result As String
    Dim strOperator, strFilter As String
    

    On Error GoTo Err_PW
    strOutput = ""
    strAuthorization = strN4Authorization ' "Basic bjRhcGk6d2VsY29tZQ=="
    strFilter = strN4Url & "/apex/api/query?filtername=GETSOCATTRIBUTES&PARM_ContNum="
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strAsmxUrl = strFilter & pCont & strOperator
    strSoapActionUrl = "POST " & "/apex/services/inventoryservice HTTP/1.1"
    
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    objDom.async = False
    objDom.loadxml strXmlBody

    objXmlHttp.Open "GET", strAsmxUrl, False, "n4api", "welcome"
    
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", strAuthorization
    
    objXmlHttp.send objDom.xml

    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    
    
    'SHARON BEGIN
'    strQuery = "//query-response/data-table"
'    result = objResult.selectSingleNode(strQuery).Text
'
'    strQuery = "//query-response/data-table/rows"
'    result = objResult.selectSingleNode(strQuery).Text
    
    'SHARON END
    
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
        
    Call ParseSOC(result, pSize, pStorageStart, pStorageEnd, bHasPaidThruDate)
    

    Set objXmlHttp = Nothing
    
 
    Sparcs_GetSOC = strRet
    Exit Function
    
Err_PW:
    Sparcs_GetSOC = "Error: " & Err.Number & " - " & Err.Description
End Function
    

Public Function ParseSOC(ByVal sQResult As String, ByRef pSize As String, ByRef pStorageStart As String, ByRef pStorageEnd As String, ByRef bHasPaidThruDate As Boolean)
'http://sbitc-dev:9080/apex/api/query?filtername=GetSOCATTRIBUTES&PARM_ContNum=EDWU7762112&operatorId=ICTSI&complexId=PH&facilityId=SBITC
'Assumption: universal query returns storage paid thru day, size, time-in
Dim arrQList() As String
Dim strElem As String
Dim iCtr As Integer
Dim iArrLen As Integer


bHasPaidThruDate = False
pSize = ""
pStorageStart = ""
pStorageEnd = ""
arrQList = Split(sQResult, " ")
iArrLen = UBound(arrQList)

If iArrLen > 0 Then
    iCtr = 0
    strElem = arrQList(iCtr)
    If IsDate(strElem) Then
       iCtr = iCtr + 1  ' iCtr = 1
       pStorageStart = "20" & strElem '& " " & arrQList(iCtr)  'Has storage paid thru date
       iCtr = iCtr + 1  ' iCtr = 2
       pSize = Left(arrQList(iCtr), 2)
       iCtr = iCtr + 1  ' iCtr = 3
       bHasPaidThruDate = True
    Else
        pSize = Left(strElem, 2)
        iCtr = iCtr + 1  ' iCtr = 1
    End If
    'If iCtr <= iArrLen Then
    If bHasPaidThruDate = False Then
        pStorageStart = "20" & arrQList(iCtr) '& " " & arrQList(iCtr + 1) ' retrieve Time-in
    End If
End If

End Function
'sharon end

'Arnel simula
Public Function ReleaseStuffing(ByVal pContNum As String)
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor, strUnit As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
'    strAuthorization = "Basic bjRhcGk6d2VsY29tZQ=="
    strUrl = strN4Url & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/argoservice HTTP/1.1"
  
    strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:arg=""http://www.navis.com/services/argoservice"" xmlns:v1=""http://types.webservice.argo.navis.com/v1.0""> " & _
"   <soapenv:Header/> " & _
"   <soapenv:Body> " & _
"      <arg:genericInvoke> " & _
"         <arg:scopeCoordinateIdsWsType> " & _
"            <!--Optional:--> " & _
"            <v1:operatorId>ICTSI</v1:operatorId> " & _
"            <!--Optional:--> " & _
"            <v1:complexId>PH</v1:complexId> " & _
"            <!--Optional:--> " & _
"            <v1:facilityId>SBITC</v1:facilityId> " & _
"            <!--Optional:--> " & _
"            <v1:yardId>SBITC</v1:yardId> " & _
"         </arg:scopeCoordinateIdsWsType> " & _
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id=""" & pContNum

    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""STUFF_PERMISSION"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & strUnit
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strN4Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function

Public Function Sparcs_LastFreeDay(ByVal pContNum As String, ByVal pCharge As String, ByVal pGKey As String, ByVal pPaid As String)
 Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
    'strAuthorization = "Basic bjRhcGk6d2VsY29tZQ==" 'c3NhbmNoZXo6cGFzc3dvcmQ="
    'strAuthorization = "Basic c3NhbmNoZXo6cGFzc3dvcmQ="
    'strUrl = "http://172.16.0.219:9080/apex/services/inventoryservice?wsdl"
    'strSoapAction = "POST http://172.16.0.219:9080/apex/services/inventoryservice HTTP/1.1"
    
    'strUrl = "http://192.168.11.151:9080/apex/services/inventoryservice?wsdl"
    ' strSoapAction = "POST http://192.168.11.151:9080/apex/services/inventoryservice HTTP/1.1"
    
    strUrl = strN4Url & "/apex/services/inventoryservice?wsdl"
    strSoapAction = "POST " & strN4Url & "/apex/services/inventoryservice HTTP/1.1"
  
  strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
"   <soapenv:Header/>" & _
"   <soapenv:Body>" & _
"      <inv:proposePaidThruDay>" & _
"        <inv:scopeCoordinateIdsWsType>" & _
"            <!--Optional:-->" & _
"            <v1:operatorId>ICTSI</v1:operatorId>" & _
"            <!--Optional:-->" & _
"            <v1:complexId>PH</v1:complexId>" & _
"            <!--Optional:-->" & _
"            <v1:facilityId>SBITC</v1:facilityId>" & _
"            <!--Optional:-->" & _
"            <v1:yardId>SBITC</v1:yardId>" & _
"         </inv:scopeCoordinateIdsWsType>" & _
"         <inv:eqId>"

strChargeFor = "</inv:eqId>" & _
"         <inv:chargeFor>"

strPaid = "</inv:chargeFor>" & _
"         <inv:paidThruDay>"

strGKey = "</inv:paidThruDay>" & _
"         <inv:extractGkey>"

strSoapEnd = "</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd


    ' Call PostWebservice and put result in text box
    If pCharge = "STORAGE" Then
        strOutput = GetLastFreeDay(strUrl, strSoapAction, strXML, strN4Authorization)
    'need to consider if will create for reefer container
'    ElseIf pCharge = "REEFER" Then
'        strOutput = GetReefer(strUrl, strSoapAction, strXML, strN4Authorization)
    End If
End Function

Public Function GetLastFreeDay(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim strQueryDays As String
    Dim result As String
    Dim resultDays As String
    Dim bError As Boolean
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    'Set currNode = CreateObject("MSXML2.XMLDOMNode")
    
    ' Load XML
    objDom.async = False
    objDom.loadxml XmlBody

    ' Open the webservice
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext

    objResult.async = False
    objResult.loadxml strRet
    'strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns5:previousPaidThruDay"
    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns3:lastFreeDay"
'    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

'Added
'    strQueryDays = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns9:daysPaid"
'    resultDays = objResult.selectSingleNode(strQueryDays).Text
'    If CInt(resultDays) > 0 Then
'        strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns5:previousPaidThruDay"
'        result = Left(objResult.selectSingleNode(strQuery).Text, 10)
'    ElseIf CInt(resultDays) = 0 Then
'        strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns3:lastFreeDay"
'        result = Left(objResult.selectSingleNode(strQuery).Text, 10)
'    End If

    'Edited by Navis Project Team 11/05/2009
'    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns2:firstFreeDay"
'    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns5:previousPaidThruDay"
    bError = IsError(objResult.selectSingleNode(strQuery).Text)
    
    If bError Then
        result = "1899-12-30 00:00:00"
    Else
        result = Left(objResult.selectSingleNode(strQuery).Text, 10)
    End If
    
    dStorage = result
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetLastFreeDay = strRet
    Exit Function
    
Err_PW:
    GetLastFreeDay = "Error: " & Err.Number & " - " & Err.Description

End Function

Public Function Sparcs_GETCONTOOG(ByVal strContNum As String)
     
    Dim strSoapAction As String
    Dim strUrl, strFilter, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    
    strOutput = ""
    strFilter = strN4Url & "/apex/api/query?filtername=GETCONTOOG&PARM_CONTNO="
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strOperator
    strSoapAction = "POST " & strN4Url & "/apex/services/inventoryservice HTTP/1.1"
 

    strOutput = GetOOG(strUrl, strSoapAction, "", strN4Authorization)
End Function

Public Sub UpdateIsN4BillingDGPermissionGranted(strContNo As String, intCCRNum As Integer)

End Sub
'Arnel tapos

