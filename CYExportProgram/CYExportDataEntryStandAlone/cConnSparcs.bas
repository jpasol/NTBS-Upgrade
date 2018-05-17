Attribute VB_Name = "cConnSparcs"
Public strN4Server As String
Public strN4Authorization As String
Public strN4UserName As String
Public strN4Password As String


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
    strAuthorization = strN4Authorization
    strUrl = strN4Server & "/apex/services/inventoryservice?wsdl"
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  
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
        strOutput = GetDischarge(strUrl, strSoapAction, strXML, strAuthorization)
    ElseIf pCharge = "REEFER" Then
        strOutput = GetReefer(strUrl, strSoapAction, strXML, strAuthorization)
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
    objXmlHttp.Open "POST", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns2:firstFreeDay"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    dStorage = result
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetDischarge = strRet
    Exit Function
    
Err_PW:
    GetDischarge = "Error: " & err.Number & " - " & err.Description

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
    objXmlHttp.Open "POST", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//soapenv:Envelope/soapenv:Body/proposePaidThruDayResponse/proposePaidThruDayResponse/ns2:outTime"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10) & " " & Mid(objResult.selectSingleNode(strQuery).Text, 12, 8)

    dReefer = result
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetReefer = strRet
    Exit Function
    
Err_PW:
    GetReefer = "Error: " & err.Number & " - " & err.Description

End Function

Public Function Sparcs_DGCode(ByVal pCont As String)
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=GETDGCODE&PARM_ContNum="
    strContNum = Trim(strParam)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    strOutput = GetDGCode(strUrl, strSoapAction, strXML, strAuthorization)
End Function

Public Function Sparcs_FreightKind(ByVal pCont As String)
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=GET_FREIGHT_KIND_EXPORT&PARM_ContNum="
    strContNum = Trim(strParam)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    strOutput = GetFreightKind(strUrl, strSoapAction, strXML, strAuthorization)

End Function

Public Function GetFreightKind(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
     Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim result As String
    Dim strFreightKind As String

    
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
    'Text1.Text = result
    
    strFreightKind = Trim(Left(result, 6))
    
    If strFreightKind = "Empty" Then
        strFreightKind = "E"
    ElseIf strFreightKind = "FCL" Then
        strFreightKind = "F"
    End If
    
    frmCCRde06.utxtFEmp.Text = strFreightKind
    
    ' Close object
    Set objXmlHttp = Nothing
     
    ' Return result
    GetFreightKind = strRet
    Exit Function
    
Err_PW:
    frmCCRde06.utxtFEmp.Text = " "
    GetFreightKind = "Error: " & err.Number & " - " & err.Description
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    GetDGCode = "Error: " & err.Number & " - " & err.Description

End Function

Public Function Sparcs_Commodity(ByVal pCont As String)
        ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=GET_COMMODITY_EXPORT&PARM_ContNum="
    strContNum = Trim(pCont)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    strOutput = GetCommodityCode(strUrl, strSoapAction, strXML, strAuthorization)
End Function

Public Function Sparcs_Shipper(ByVal pCont As String)
        ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=GET_SHIPPER_EXPORT&PARM_ContNum="
    strContNum = Trim(pCont)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    strOutput = GetShipper(strUrl, strSoapAction, strXML, strAuthorization)
End Function

Public Function GetShipper(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
        
    frmCCRde06.utxtExporter.Text = result
    
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetShipper = strRet
    Exit Function
    
Err_PW:
    GetShipper = "Error: " & err.Number & " - " & err.Description
    
End Function

Public Function Sparcs_VisitReference(ByVal pCont As String) As String
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
        
    strParam = pCont
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=GET_CONT_VISIT_REF_EXPORT&PARM_ContNum="
    strContNum = Trim(pCont)
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    Sparcs_VisitReference = GetVisitReference(strUrl, strSoapAction, strXML, strAuthorization)
End Function

Public Function GetVisitReference(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
        
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetVisitReference = result
    Exit Function
    
Err_PW:
    GetVisitReference = "Error: " & err.Number & " - " & err.Description
    
End Function

Public Function Sparcs_OOG(ByVal pCont As String)
     Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strAuthorization = strN4Authorization
    'strUrl = "http://172.16.0.219:9080/apex/api/query?filtername=GETOOG&PARM_ContNum=SHAR1809099&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strFilter = strN4Server & "/apex/api/query?filtername=GETOOG&PARM_ContNum="
    strContNum = Trim(strParam)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  
    strOutput = GetOOG(strUrl, strSoapAction, "", strAuthorization)
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

    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    GetOOG = "Error: " & err.Number & " - " & err.Description
End Function

Public Function ParseOOG(ByVal sName As String)
Dim sLetter As String
Dim iAsc As Integer
Dim i As Integer
Dim ctrInt, ctrNo As Integer
Dim iNum1, iNum2 As Integer

ctrInt = 1
ctrNo = 1

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
                'frmManifestCont.mskOVLength.Text = iNum1 + iNum2
                frmCCRde06.utxtLength.Value = iNum1 + iNum2
                iNum1 = Empty
                iNum2 = Empty
                ctrNo = ctrNo + 1
                ctrInt = 1
                sName = Replace(sName, Left(sName, 1), "", 1, 1)
            ElseIf ctrNo = 2 Then
                'frmManifestCont.mskOVWidth.Text = iNum1 + iNum2
                frmCCRde06.utxtWidth.Value = iNum1 + iNum2
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
frmCCRde06.utxtHeight.Value = iNum1 + iNum2
iNum1 = Empty
iNum2 = Empty
ctrNo = ctrNo + 1
ctrInt = 1
sName = Replace(sName, Left(sName, 1), "", 1, 1)


End Function

'Added by Navis Project Team 10/26/2009
Public Function ParseDG(ByVal sName As String)
Dim strDG As String
Dim arrDGList() As String
Dim arrDGVal() As String
Dim iCtrDG As Integer
Dim iCtrDGVal As Integer
Dim intHighestDG As Integer
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
'            arrDGVal = Split(arrDGList(iCtrDG), ".")
'            intHighestDG = CInt(arrDGVal(0))
'        End If
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


Select Case intHighestDG
    Case 0
        frmCCRde06.utxtNumDangr.Value = " " & Chr(124) & " Not Applicable"
    Case 1
        frmCCRde06.utxtNumDangr.Value = "1" & Chr(124) & " Explosives DC1"
    Case 2
        frmCCRde06.utxtNumDangr.Value = "2" & Chr(124) & " Gases DC2"
    Case 3
        frmCCRde06.utxtNumDangr.Value = intHighestDG
    Case 4
        frmCCRde06.utxtNumDangr.Value = "4" & Chr(124) & " Inflammable Solids DC2 "
    Case 5
        frmCCRde06.utxtNumDangr.Value = "5" & Chr(124) & " Oxidizing Agents/Organic Peroxides DC3"
    Case 6
        frmCCRde06.utxtNumDangr.Value = "6" & Chr(124) & " Poisonous(toxic) and Infectious Substances DC1"
    Case 7
        frmCCRde06.utxtNumDangr.Value = "7" & Chr(124) & " Radioactive Substances DC2"
    Case 8
        frmCCRde06.utxtNumDangr.Value = "8" & Chr(124) & " Corrosives DC1"
    Case 9
        frmCCRde06.utxtNumDangr.Value = "9" & Chr(124) & " Miscellaneous Dangerous Substances DC3"
End Select

End Function

'Public Function ParseDG(ByVal sName As String)
'Dim sLetter As String
'Dim iAsc As Integer
'Dim i As Integer
'Dim ctrInt, ctrNo As Integer
'Dim strNum1, strNum2, strNum3 As String
'Dim length, width, height As Integer
'Dim DGCode As Integer
'
'ctrInt = 1
'ctrNo = 1
'strNum1 = "0"
'strNum2 = "0"
'strNum3 = "0"
'For i = 1 To Len(sName)
'    sLetter = Mid(sName, 1, 1)
'    iAsc = Asc(sLetter)
'
'    If (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 122) Or iAsc = 46 Then
'            If ctrInt = 1 Then
'                strNum1 = ""
'                strNum1 = strNum1 & Left(sName, 1)
'                sName = Replace(sName, Left(sName, 1), "", 1, 1)
'            ElseIf ctrInt = 2 Then
'                strNum2 = ""
'                strNum2 = strNum2 & Left(sName, 1)
'                sName = Replace(sName, Left(sName, 1), "", 1, 1)
'            ElseIf ctrInt = 3 Then
'                strNum3 = ""
'                strNum3 = strNum3 & Left(sName, 1)
'                sName = Replace(sName, Left(sName, 1), "", 1, 1)
'            End If
'    Else
'        If iAsc = 44 Then
'            sName = Replace(sName, Left(sName, 1), "", 1, 1)
'            ctrInt = ctrInt + 1
'        End If
'    End If
'
'Next
'
''If CInt(Left(strNum1, 1)) > CInt(Left(strNum2, 1)) And CInt(Left(strNum1, 1)) > CInt(Left(strNum3, 1)) Then
''    DGCode = CInt(Left(strNum1, 1))
''ElseIf CInt(Left(strNum2, 1)) > CInt(Left(strNum1, 1)) And CInt(Left(strNum2, 1)) > CInt(Left(strNum3, 1)) Then
''    DGCode = CInt(Left(strNum2, 1))
''ElseIf CInt(Left(strNum3, 1)) > CInt(Left(strNum1, 1)) And CInt(Left(strNum3, 1)) > CInt(Left(strNum2, 1)) Then
''    DGCode = CInt(Left(strNum3, 1))
''End If
'
'If IsNumeric(strNum1) And IsNumeric(strNum2) And IsNumeric(strNum3) Then
'    If CInt(Left(strNum1, 1)) > CInt(Left(strNum2, 1)) And CInt(Left(strNum1, 1)) > CInt(Left(strNum3, 1)) Then
'        DGCode = CInt(Left(strNum1, 1))
'    ElseIf CInt(Left(strNum2, 1)) > CInt(Left(strNum1, 1)) And CInt(Left(strNum2, 1)) > CInt(Left(strNum3, 1)) Then
'        DGCode = CInt(Left(strNum2, 1))
'    ElseIf CInt(Left(strNum3, 1)) > CInt(Left(strNum1, 1)) And CInt(Left(strNum3, 1)) > CInt(Left(strNum2, 1)) Then
'        DGCode = CInt(Left(strNum3, 1))
'    End If
'Else
'    If IsNumeric(strNum1) And (IsEmpty(strNum2) Or IsNumeric(strNum2) = False) And (IsEmpty(strNum3) Or IsNumeric(strNum3) = False) Then
'        DGCode = strNum1
'    ElseIf IsNumeric(strNum1) And IsNumeric(strNum2) = True And (IsEmpty(strNum3) Or IsNumeric(strNum3) = False) Then
'        If CInt(Left(strNum1, 1)) > CInt(Left(strNum2, 1)) Then
'            DGCode = strNum1
'        ElseIf CInt(Left(strNum1, 1)) < CInt(Left(strNum2, 1)) Then
'            DGCode = strNum2
'        End If
'    End If
'End If
'
'If DGCode = 0 Then
'    frmCCRde06.utxtNumDangr.Value = " " & Chr(124) & " Not Applicable"
'ElseIf DGCode = 1 Then
'    frmCCRde06.utxtNumDangr.Value = "1" & Chr(124) & " Explosives DC1"
'ElseIf DGCode = 2 Then
'    frmCCRde06.utxtNumDangr.Value = "2" & Chr(124) & " Gases DC2"
'ElseIf DGCode = 3 Then
'    frmCCRde06.utxtNumDangr.Value = DGCode
'ElseIf DGCode = 4 Then
'    frmCCRde06.utxtNumDangr.Value = "4" & Chr(124) & " Inflammable Solids DC2 "
'ElseIf DGCode = 5 Then
'    frmCCRde06.utxtNumDangr.Value = "5" & Chr(124) & " Oxidizing Agents/Organic Peroxides DC3"
'ElseIf DGCode = 6 Then
'    frmCCRde06.utxtNumDangr.Value = "6" & Chr(124) & " Poisonous(toxic) and Infectious Substances DC1"
'ElseIf DGCode = 7 Then
'    frmCCRde06.utxtNumDangr.Value = "7" & Chr(124) & " Radioactive Substances DC2"
'ElseIf DGCode = 8 Then
'    frmCCRde06.utxtNumDangr.Value = "8" & Chr(124) & " Corrosives DC1"
'ElseIf DGCode = 9 Then
'    frmCCRde06.utxtNumDangr.Value = "9" & Chr(124) & " Miscellaneous Dangerous Substances DC3"
'End If
'End Function

Public Function GetGKey(ByVal pCont As String, ByVal pStatus As String, ByVal pType As String) As String
    Dim rstGKey As ADODB.Recordset
    Dim strGKey As String
    
    Set rstGKey = New ADODB.Recordset
    
    strGKey = ""
    strGKey = "SELECT gkey FROM argo_chargeable_unit_events " & _
                "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' and event_type = '" & pType & "'"

    rstGKey.Open strGKey, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    GetGKey = rstGKey.Fields(0)
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
    strAuthorization = strN4Authorization
    strUrl = strN4Server & "/apex/services/inventoryservice?wsdl"
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  
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

strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd


    ' Call PostWebservice and put result in text box
    If pCharge = "STORAGE" Then
        strOutput = SaveStorage(strUrl, strSoapAction, strXML, strAuthorization)
    ElseIf pCharge = "REEFER" Then
        strOutput = GetReefer(strUrl, strSoapAction, strXML, strAuthorization)
    End If
End Function

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
    objXmlHttp.Open "POST", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    dStorage = result
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    SaveStorage = strRet
    Exit Function
    
Err_PW:
    SaveStorage = "Error: " & err.Number & " - " & err.Description

End Function

Public Function Sparcs_ExpSize(ByVal pCont As String)
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont '"SHAR1809099"
    strOutput = ""
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=GETEXPSZE&PARM_ContNum="
    strContNum = Trim(strParam)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    strOutput = GetExpSze(strUrl, strSoapAction, strXML, strAuthorization)
End Function

Public Function GetExpSze(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
        
    frmCCRde06.utxtSze.Value = Left(result, 2)
    
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetExpSze = strRet
    Exit Function
    
Err_PW:
    GetExpSze = "Error: " & err.Number & " - " & err.Description

End Function

Public Function Sparcs_VesselName(ByVal pCont As String)
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSoapAction As String
    Dim strUrl, strFilter, strVisitRef, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam, strParamCat As String
    Dim strOutput As String
    
    strParam = pCont '"SHAR1809099"
    strOutput = ""
    strAuthorization = strN4Authorization
    strFilter = strN4Server & "/apex/api/query?filtername=VESSEL_NAME&PARM_VisitReference="
    strVisitRef = Sparcs_VisitReference(strParam)
    strCategory = "&PARM_Category="
    strParamCat = "EXPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strVisitRef & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

    strXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
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
"         <inv:eqId>APHU6623529</inv:eqId>" & _
"         <inv:chargeFor>STORAGE</inv:chargeFor>" & _
"         <inv:paidThruDay>?</inv:paidThruDay>" & _
"         <inv:extractGkey>4992</inv:extractGkey>" & _
"      </inv:proposePaidThruDay>" & _
"   </soapenv:Body>" & _
"</soapenv:Envelope>"

    ' Call PostWebservice and put result in text box
    strOutput = GetVesselName(strUrl, strSoapAction, strXML, strAuthorization)
End Function

Public Function GetVesselName(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
    frmCCRde06.utxtVessel.Text = result
    
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetVesselName = strRet
    Exit Function
    
Err_PW:
    GetVesselName = "Error: " & err.Number & " - " & err.Description

End Function

Public Function GetCommodityCode(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
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
    objXmlHttp.Open "GET", AsmxUrl, False, "strN4UserName", "strN4Password"
    
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
    strQuery = "//query-response/data-table/rows/row"
    result = objResult.selectSingleNode(strQuery).Text
        
    frmCCRde06.utxtCommodity.Text = result
    
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetCommodityCode = strRet
    Exit Function
    
Err_PW:
    GetCommodityCode = "Error: " & err.Number & " - " & err.Description

End Function

Public Function ReleaseHold(ByVal pContNum As String)
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
    Dim strQuery As String
    strOutput = ""
    strAuthorization = strN4Authorization
    strUrl = strN4Server & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Server & "/apex/services/argoservice HTTP/1.1"
  
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
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id="""

    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & pContNum & strUnit
'Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objResult = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, "strN4UserName", "strN4Password"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strAuthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
    
    ''
        ' Get all response text from webservice
    objResult.async = False
    objResult.loadxml strParam
    strQuery = "//genericInvokeResponse/ns1:commonResponse/ns1:Status"
    ReleaseHold = objResult.selectSingleNode(strQuery).Text
    ''
    
   'Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function

'added Navis Project Team 10/26/2009
'Public Function ReleaseDGHold(ByVal pContNum As String)
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
'    Dim strQuery As String
'    strOutput = ""
'    strAuthorization = strN4Authorization
'    strUrl = "http://sbitc-dev:9080/apex/services/argoservice?wsdl"
'    strSoapAction = "POST http://sbitc-dev:9080/apex/services/argoservice HTTP/1.1"
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
'"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id="""
'
'    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING_DG"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
'"      </arg:genericInvoke> " & _
'"   </soapenv:Body> " & _
'"</soapenv:Envelope>"
'
'    strXML = strScope & pContNum & strUnit
''Debug.Print strXML
'
'    Set objDom = CreateObject("MSXML2.DOMDocument")
'    Set objResult = CreateObject("MSXML2.DOMDocument")
'    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
'
'
'    ' Load XML
'    objDom.async = False
'    objDom.loadxml strXML
'
'    ' Open the webservice
'    objXmlHttp.Open "POST", strUrl, False, "strN4UserName", "strN4Password"
'
'    ' Create headings
'    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
'    objXmlHttp.setRequestHeader "Authorization", strAuthorization
'
'    ' Send XML command
'    objXmlHttp.send objDom.xml
'    strParam = objXmlHttp.responsetext
'
'    ''
'        ' Get all response text from webservice
'    objResult.async = False
'    objResult.loadxml strParam
'    strQuery = "//genericInvokeResponse/ns1:commonResponse/ns1:Status"
'    ReleaseDGHold = objResult.selectSingleNode(strQuery).Text
'    ''
'
'   'Debug.Print strParam
'    ' Close object
'    Set objXmlHttp = Nothing
'
'End Function
Public Function ReleaseDGHold(ByVal pContNum As String)
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
    strAuthorization = strN4Authorization
    strUrl = strN4Server & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Server & "/apex/services/argoservice HTTP/1.1"
  
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
    objXmlHttp.Open "POST", strUrl, False, "strN4UserName", "strN4Password"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strAuthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strRet = objXmlHttp.responsetext
    objResult.async = False
    objResult.loadxml strRet
    
    strQuery = "//soapenv:Envelope/soapenv:Body/genericInvokeResponse/genericInvokeResponse/ns1:commonResponse/ns1:Status"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    ' Close object
    Set objXmlHttp = Nothing
    ReleaseDGHold = result
    Exit Function
ERR_Handler:
    ReleaseDGHold = "Error: " & err.Number & " - " & err.Description
End Function



'Public Function ReleaseOOGHold(ByVal pContNum As String)
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
'    Dim strQuery As String
'    strOutput = ""
'    strAuthorization = strN4Authorization
'    strUrl = "http://sbitc-dev:9080/apex/services/argoservice?wsdl"
'    strSoapAction = "POST http://sbitc-dev:9080/apex/services/argoservice HTTP/1.1"
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
'"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit id="""
'
'    strUnit = """></unit></units></entities><flags><flag hold-perm-id=""BILLING_OOG"" action=""GRANT_PERMISSION""/></flags></hpu>]]></arg:xmlDoc> " & _
'"      </arg:genericInvoke> " & _
'"   </soapenv:Body> " & _
'"</soapenv:Envelope>"
'
'    strXML = strScope & pContNum & strUnit
''Debug.Print strXML
'
'    Set objDom = CreateObject("MSXML2.DOMDocument")
'    Set objResult = CreateObject("MSXML2.DOMDocument")
'    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
'
'
'    ' Load XML
'    objDom.async = False
'    objDom.loadxml strXML
'
'    ' Open the webservice
'    objXmlHttp.Open "POST", strUrl, False, "strN4UserName", "strN4Password"
'
'    ' Create headings
'    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
'    objXmlHttp.setRequestHeader "Authorization", strAuthorization
'
'    ' Send XML command
'    objXmlHttp.send objDom.xml
'    strParam = objXmlHttp.responsetext
'
'    ''
'        ' Get all response text from webservice
'    objResult.async = False
'    objResult.loadxml strParam
'    strQuery = "//genericInvokeResponse/ns1:commonResponse/ns1:Status"
'    ReleaseOOGHold = objResult.selectSingleNode(strQuery).Text
'    ''
'
'   'Debug.Print strParam
'    ' Close object
'    Set objXmlHttp = Nothing
'
'End Function
Public Function ReleaseOOGHold(ByVal pContNum As String)
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
    strAuthorization = strN4Authorization
    strUrl = strN4Server & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & strN4Server & "/apex/services/argoservice HTTP/1.1"
  
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
    objXmlHttp.Open "POST", strUrl, False, "strN4UserName", "strN4Password"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strAuthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strRet = objXmlHttp.responsetext
    objResult.async = False
    objResult.loadxml strRet
    
    strQuery = "//soapenv:Envelope/soapenv:Body/genericInvokeResponse/genericInvokeResponse/ns1:commonResponse/ns1:Status"
    result = Left(objResult.selectSingleNode(strQuery).Text, 10)

    ' Close object
    Set objXmlHttp = Nothing
    ReleaseOOGHold = result
    Exit Function
ERR_Handler:
    ReleaseOOGHold = "Error: " & err.Number & " - " & err.Description
End Function

Public Function GKeyHasUnitOut(Unit_GKey As String) As Boolean
    Dim rsGKey As ADODB.Recordset
    Dim strQuery As String
    Dim bResult As Boolean

    Set rsGKey = New ADODB.Recordset

    bResult = False
    
    strQuery = "SELECT unit_gkey FROM argo_chargeable_unit_events WHERE unit_gkey='" & Unit_GKey & "' AND event_type LIKE 'UNIT_OUT%' AND status <> 'CANCELLED' "
    
    rsGKey.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    If Not rsGKey.BOF Or Not rsGKey.EOF Then
        bResult = True
    End If
    
    rsGKey.Close
    Set rsGKey = Nothing
    
    GKeyHasUnitOut = bResult
End Function

Public Sub ReadConfig()
Dim Xcnt As Integer
 
Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1

Do While Not EOF(1)
    Xcnt = Xcnt + 1
    Select Case Xcnt
        Case 1
            Line Input #1, sqlConBilling
        Case 2
            Line Input #1, sqlConNavis
    End Select
Loop
End Sub

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
