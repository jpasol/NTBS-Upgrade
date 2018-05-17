Attribute VB_Name = "cConnSparcs"

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
    'for dev
    'strAuthorization = "Basic bjRhcGk6d2VsY29tZQ=="
    'for production
    strauthorization = strN4Authorization '"Basic bjRhcGk6d2VsY29tZQ=="
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
        strOutput = GetDischarge(strUrl, strSoapAction, strXML, strauthorization)
    ElseIf pCharge = "REEFER" Then
        strOutput = GetReefer(strUrl, strSoapAction, strXML, strauthorization)
    End If
End Function


'Public Function Sparcs_Reefer()
' Dim strSoapAction As String
'    Dim strUrl As String
'    Dim strXML As String
'    Dim strParam As String
'    Dim strOutput As String
'    Dim strScope As String
'    Dim strChargeFor As String
'    Dim strPaid As String
'    Dim strGKey As String
'    Dim strSoapEnd As String
'    strOutput = ""
'    'for dev
'    'strAuthorization = "Basic bjRhcGk6d2VsY29tZQ=="
'    'for production
'    strauthorization = strN4Authorization '"Basic bjRhcGk6d2VsY29tZQ=="
'    strUrl = strN4Server & "/apex/services/argoservice?wsdl"
'    strSoapAction = "POST " & strN4Server & "/apex/services/argoservice HTTP/1.1"   '"POST http://sbitc-dev:9080/apex/services/inventoryservice HTTP/1.1"
'
'  strScope = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inv=""http://www.navis.com/services/inventoryservice"" xmlns:v1=""http://types.webservice.inventory.navis.com/v1.0"">" & _
'"   <soapenv:Header/>" & _
'"   <soapenv:Body>" & _
'"      <arg:genericInvoke>" & _
'"        <arg:scopeCoordinateIdsWsType>" & _
'"            <!--Optional:-->" & _
'"            <v1:operatorId>ICTSI</v1:operatorId>" & _
'"            <!--Optional:-->" & _
'"            <v1:complexId>PH</v1:complexId>" & _
'"            <!--Optional:-->" & _
'"            <v1:facilityId>SBITC</v1:facilityId>" & _
'"            <!--Optional:-->" & _
'"            <v1:yardId>SBITC</v1:yardId>" & _
'"         </arg:scopeCoordinateIdsWsType>" & _
'"         <arg:xmlDoc>"
'
'strChargeFor = "<![CDATA[<groovy class-location=""database"" class-name=""getPowerConnectPaidThruTime"" > """
'
'strPaid = "<parameters><parameter id=""equipment-id"" value=""EDWU4409129""/></parameters>"
'
'strGKey = "</groovy>]]></arg:xmlDoc></arg:genericInvoke>"
'
'strSoapEnd = "</soapenv:Body>" & _
'"</soapenv:Envelope>"
'
'strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd
'
'
'    ' Call PostWebservice and put result in text box
'
'        strOutput = GetReefer(strUrl, strSoapAction, strXML, strauthorization)
'End Function

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
    objXmlHttp.Open "POST", AsmxUrl, False, strN4UserName, strN4Password    '"n4api", "welcome"
    
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

Public Sub Sparcs_Reefer(ByVal pContNum As String)
'PRNH
'    Dim rst As ADODB.Recordset
'    Dim query As String
'
'    Set rst = New ADODB.Recordset
'    query = "select top 1 flex_date01 from argo_chargeable_unit_events " & _
'        "where equipment_id = '" & pContNum & "' and event_type = 'UNIT_POWER_CONNECT' " & _
'        "and ufv_time_out is null order by flex_date01 desc"
'
'    rst.Open query, gcnnNavis, adOpenForwardOnly, adLockReadOnly
'
'    If Not rst.BOF Or Not rst.EOF Then
'        dReefer = rst.Fields(0)
'    Else
'        dReefer = "1899-12-30 00:00:00"
'    End If
'
'    Set rst = Nothing


    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    Dim strScope As String
    Dim strChargeFor, strUnit, strEnd As String
    Dim strPaid As String
    Dim strGKey As String
    Dim strSoapEnd As String
    strOutput = ""
    strauthorization = strN4Authorization '"Basic bjRhcGk6d2VsY29tZQ=="
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
"         <arg:xmlDoc><![CDATA[<groovy class-location=""database"" class-name=""getPowerConnectPaidThruTime"

    strUnit = """><parameters><parameter id=""equipment-id"" value="""

    strEnd = """/></parameters></groovy>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope & strUnit & pContNum & strEnd
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")


    ' Load XML
    objDom.async = False
    objDom.loadxml strXML

    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password '"n4api", "welcome"

    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strauthorization

    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
   strOutput = GetReefer(strUrl, strSoapAction, strParam, strauthorization)
    ' Close object
    Set objXmlHttp = Nothing

    ' Call PostWebservice and put result in text box
        
End Sub

Public Function GetReefer(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim objResult As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strQuery As String
    Dim strQueryPaidThru  As String
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
    objXmlHttp.Open "POST", AsmxUrl, False, "n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.setRequestHeader "Authorization", Authorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responsetext
    'MsgBox Mid(XmlBody, 700, 37)
    objResult.async = False
    objResult.loadxml strRet
    'strQuery = "//soapenv:Envelope/soapenv:Body/genericInvokeResponse/genericInvokeResponse/ns1:commonResponse/ns1:MessageCollector/ns1:Messages/Message"
    
    Dim nPaidThruPos As Integer
    Dim bHasPaidThru As Boolean
    Dim bHasPlugIn As Boolean
    
    bHasPaidThru = False
    bHasPlugIn = False
    
    nPaidThruPos = InStr(XmlBody, "PowerPaidThruTime:")
    If nPaidThruPos > 0 Then
        strQueryPaidThru = Mid(XmlBody, nPaidThruPos)
        If IsDate(Mid(strQueryPaidThru, 19, 19)) Then
            result = Mid(strQueryPaidThru, 19, 19) 'objResult.selectSingleNode(strQuery).Text
            dReefer = result
            bHasPaidThru = True
        End If
    End If
   If bHasPaidThru = False Then
        strQuery = Mid(XmlBody, 701, 37)
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
    
    
    'MsgBox currNode.Text
    'Text1.Text = currNode
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
    GetReefer = strRet
    Exit Function
    
Err_PW:
    GetReefer = "Error: " & err.Number & " - " & err.Description
MsgBox err.Description
End Function

Public Function Sparcs_DGCode(ByVal pCont As String, ByVal strParamCat As String)
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
    Dim strParam As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strauthorization = strN4Authorization   '"Basic bjRhcGk6d2VsY29tZQ=="
    strFilter = strN4Server & "/apex/api/query?filtername=GETDGCODE&PARM_ContNum="
    strContNum = strParam
    strCategory = "&PARM_Category="
    'strParamCat = "IMPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & strN4Server & "/apex/services/inventoryservice HTTP/1.1"
  

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
'
'    ' Call PostWebservice and put result in text box
'    strOutput = GetDGCode(strUrl, strSoapAction, strXML, strauthorization)
strOutput = GetDGCode(strUrl, strSoapAction, "", strauthorization)
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
    objXmlHttp.Open "GET", AsmxUrl, False, strN4UserName, strN4Password '"n4api", "welcome"
    
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

Public Function Sparcs_OOG(ByVal pCont As String, ByVal strParamCat As String)
     Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strauthorization = "Basic bjRhcGk6d2VsY29tZQ=="
    'strUrl = "http://172.16.0.219:9080/apex/api/query?filtername=GETOOG&PARM_ContNum=SHAR1809099&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strFilter = strN4Server & "/apex/api/query?filtername=GETOOG&PARM_ContNum="  '"http://sbitc-dev:9080/apex/api/query?filtername=GETOOG&PARM_ContNum="
    strContNum = strParam
    strCategory = "&PARM_Category="
    strParamCat = "IMPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    strSoapAction = "POST " & "/apex/services/inventoryservice HTTP/1.1"
  

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

    'strOutput = GetOOG(strUrl, strSoapAction, strXML, strauthorization)
    strOutput = GetOOG(strUrl, strSoapAction, "", strauthorization)
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

    objXmlHttp.Open "GET", AsmxUrl, False, strN4UserName, strN4Password '"n4api", "welcome"
    
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
                frmManifestCont.mskOVLength.Text = iNum1 + iNum2
                iNum1 = Empty
                iNum2 = Empty
                ctrNo = ctrNo + 1
                ctrInt = 1
                sName = Replace(sName, Left(sName, 1), "", 1, 1)
            ElseIf ctrNo = 2 Then
                frmManifestCont.mskOVWidth.Text = iNum1 + iNum2
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
frmManifestCont.mskOVHeight.Text = iNum1 + iNum2
iNum1 = Empty
iNum2 = Empty
ctrNo = ctrNo + 1
ctrInt = 1
sName = Replace(sName, Left(sName, 1), "", 1, 1)
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
'
'For i = 1 To Len(sName)
'    sLetter = Mid(sName, 1, 1)
'    iAsc = Asc(sLetter)
'
'    If (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 122) Or iAsc = 46 Then
'            If ctrInt = 1 Then
'                strNum1 = strNum1 & Left(sName, 1)
'                sName = Replace(sName, Left(sName, 1), "", 1, 1)
'            ElseIf ctrInt = 2 Then
'                strNum2 = strNum2 & Left(sName, 1)
'                sName = Replace(sName, Left(sName, 1), "", 1, 1)
'            ElseIf ctrInt = 3 Then
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
'
'If DGCode = 0 Then
'    frmManifestCont.cboDangClass.Text = " " & Chr(124) & " Not Applicable"
'ElseIf DGCode = 1 Then
'    frmManifestCont.cboDangClass.Text = "1" & Chr(124) & " Explosives DC1"
'ElseIf DGCode = 2 Then
'    frmManifestCont.cboDangClass.Text = "2" & Chr(124) & " Gases DC2"
'ElseIf DGCode = 3 Then
'    frmManifestCont.cboDangClass.Text = "3" & Chr(124) & " Inflammable Liquid DC2"
'ElseIf DGCode = 4 Then
'    frmManifestCont.cboDangClass.Text = "4" & Chr(124) & " Inflammable Solids DC2 "
'ElseIf DGCode = 5 Then
'    frmManifestCont.cboDangClass.Text = "5" & Chr(124) & " Oxidizing Agents/Organic Peroxides DC3"
'ElseIf DGCode = 6 Then
'    frmManifestCont.cboDangClass.Text = "6" & Chr(124) & " Poisonous(toxic) and Infectious Substances DC1"
'ElseIf DGCode = 7 Then
'    frmManifestCont.cboDangClass.Text = "7" & Chr(124) & " Radioactive Substances DC2"
'ElseIf DGCode = 8 Then
'    frmManifestCont.cboDangClass.Text = "8" & Chr(124) & " Corrosives DC1"
'ElseIf DGCode = 9 Then
'    frmManifestCont.cboDangClass.Text = "9" & Chr(124) & " Miscellaneous Dangerous Substances DC3"
'End If
'End Function
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
        frmManifestCont.cboDangClass.Text = " " & Chr(124) & " Not Applicable"
    Case 1
        frmManifestCont.cboDangClass.Text = "1" & Chr(124) & " Explosives DC1"
    Case 2
        frmManifestCont.cboDangClass.Text = "2" & Chr(124) & " Gases DC2"
    Case 3
        frmManifestCont.cboDangClass.Text = "3" & Chr(124) & " Inflammable Liquid DC2"
    Case 4
        frmManifestCont.cboDangClass.Text = "4" & Chr(124) & " Inflammable Solids DC2 "
    Case 5
        frmManifestCont.cboDangClass.Text = "5" & Chr(124) & " Oxidizing Agents/Organic Peroxides DC3"
    Case 6
        frmManifestCont.cboDangClass.Text = "6" & Chr(124) & " Poisonous(toxic) and Infectious Substances DC1"
    Case 7
        frmManifestCont.cboDangClass.Text = "7" & Chr(124) & " Radioactive Substances DC2"
    Case 8
        frmManifestCont.cboDangClass.Text = "8" & Chr(124) & " Corrosives DC1"
    Case 9
        frmManifestCont.cboDangClass.Text = "9" & Chr(124) & " Miscellaneous Dangerous Substances DC3"
End Select

End Function

Public Function Sparcs_VisitID(ByVal pCont As String, ByVal strParamCat As String)
    Dim strSoapAction As String
    Dim strUrl, strFilter, strContNum, strOperator, strCategory As String
    Dim strXML As String
    Dim strParam As String
    Dim strOutput As String
    
    strParam = pCont
    strOutput = ""
    strauthorization = strN4Authorization '"Basic bjRhcGk6d2VsY29tZQ=="
    strFilter = strN4Server & "/apex/api/query?filtername=VESSEL_VISIT&PARM_Vessel Registry Number="
    strContNum = strParam
    
    'Removed by PRNH - No "Category" parameter indicated in filter--------
    'strCategory = "&PARM_Category="
    '----------------------------------------------------------------------
    
    'strParamCat = "IMPRT"
    strOperator = "&operatorId=ICTSI&complexId=PH&facilityId=SBITC&yardId=SBITC"
    
    'OLD -PRNH------------------------------------------------------------------
    'strUrl = strFilter & strContNum & strCategory & strParamCat & strOperator
    
    'PRNH - Removed the strParamCat and strCategory
    strUrl = strFilter & strContNum & strOperator
    
    
    strSoapAction = "POST " & "/apex/services/inventoryservice HTTP/1.1"
  

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
    'strOutput = GetVisitID(strUrl, strSoapAction, strXML, strAuthorization)
    
    Sparcs_VisitID = GetVisitID(strUrl, strSoapAction, strXML, strauthorization)
    
    'PRNH - If no Visit ID was retrieved
    If Sparcs_VisitID = "" Then Sparcs_VisitID = pCont
End Function

Public Function GetVisitID(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String, ByVal Authorization As String) As String
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
    objXmlHttp.Open "GET", AsmxUrl, False, strN4UserName, strN4Password '"n4api", "welcome"
    
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
    GetVisitID = result
    ' Close object
    Set objXmlHttp = Nothing
    
 
    ' Return result
'    GetVisitID = strRet
    Exit Function
    
Err_PW:
    'OLD - PRNH
    'GetVisitID = "Error: " & Err.Number & " - " & Err.Description
    
    'If no Visit ID used
    GetVisitID = ""
    MsgBox "Error in retrieving Visit ID. Kindly verify with Operations. Error message: " & err.Description

End Function

Public Function GetGKey(ByVal pCont As String, ByVal pStatus As String, ByVal pType As String, ByVal pCat As String, ByVal pVisit As String) As String
    Dim rstGKey As ADODB.Recordset
    Dim strGKey As String
    
    Set rstGKey = New ADODB.Recordset
    
    strGKey = ""
    strGKey = "SELECT gkey FROM argo_chargeable_unit_events " & _
                "Where equipment_id = '" & pCont & "' and status = '" & pStatus & "' and " & _
                "event_type = '" & pType & "' and " & _
                "category = '" & pCat & "' and ib_id= '" & pVisit & "'"

    rstGKey.Open strGKey, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    If Not rstGKey.BOF = True Or Not rstGKey.EOF = True Then
        GetGKey = rstGKey.Fields(0)
    End If
End Function

'PRNH - LDD via direct query
Public Function getLDD(ByVal regNum As String) As String
    Dim rstLDD As ADODB.Recordset
    Dim strLDD As String
    
    On Error GoTo err
    Set rstLDD = New ADODB.Recordset
    
    strLDD = "SET NOCOUNT ON; SELECT top 1 avd.time_discharge_complete " & _
        "FROM argo_carrier_visit acv " & _
        "INNER JOIN vsl_vessel_visit_details vvv ON acv.cvcvd_gkey = vvv.vvd_gkey " & _
        "INNER JOIN vsl_vessels AS vv ON vvv.vessel_gkey = vv.gkey " & _
        "INNER JOIN ref_bizunit_scoped AS rbs ON acv.operator_gkey = rbs.gkey " & _
        "INNER JOIN argo_visit_details AS avd ON acv.cvcvd_gkey = avd.gkey " & _
        "INNER JOIN ref_carrier_service AS rcs ON avd.service = rcs.gkey " & _
        "INNER JOIN inv_unit_fcy_visit AS iufv ON iufv.actual_ib_cv = acv.gkey " & _
        "INNER JOIN inv_unit AS iu ON iu.gkey = iufv.unit_gkey " & _
        "INNER JOIN ref_equipment AS req ON iu.id = req.id_full " & _
        "WHERE vvv.flex_string01 = '" & regNum & "'"

    rstLDD.Open strLDD, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    If Not rstLDD.BOF = True Or Not rstLDD.EOF = True Or Not IsNull(rstLDD.Fields(0)) Then
        getLDD = rstLDD.Fields(0)
    Else
        getLDD = "1899-12-30"
        MsgBox "No Last Discharge retrieved. Verify with Operations"
    End If
    
Exit Function
err:

MsgBox "Error retrieving LDD. Error message:" & err.Description
End Function

'PRNH - Get Company Code using RegNum (direct query)
Public Function GetCompanyCode(ByVal regNum As String) As String
    Dim rstCC As ADODB.Recordset
    Dim strQuery As String
    
    On Error GoTo err
    Set rstCC = New ADODB.Recordset
    
    strQuery = "SET NOCOUNT ON; Select vvv.flex_string02 " & _
        "from argo_carrier_visit acv INNER JOIN vsl_vessel_visit_details vvv on acv.cvcvd_gkey = vvv.vvd_gkey " & _
        "where vvv.flex_string01 = '" & regNum & "'"

    rstCC.Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly
    
    If Not rstCC.BOF = True Or Not rstCC.EOF = True Or Not IsNull(rstCC.Fields(0)) Then
        GetCompanyCode = rstCC.Fields(0)
    Else
        GetCompanyCode = ""
        MsgBox ("Company Code not indicated. Kindly verify with Operations.")
    End If
    
Exit Function
err:

MsgBox "Error retrieving LDD. Error message:" & err.Description
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
    strauthorization = strN4Authorization    '"Basic bjRhcGk6d2VsY29tZQ=="
    'strUrl = strN4Server & "/apex/services/inventoryservice HTTP/1.1"
    strUrl = strN4Server & "/apex/services/inventoryservice"
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

    ' Call PostWebservice and put result in text box
    If pCharge = "STORAGE" Then
        pPaid = Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss") & " +0800"
        strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd
        Debug.Print strXML
        strOutput = SavePaidThruDay(strUrl, strSoapAction, strXML, strauthorization)
    ElseIf pCharge = "REEFER" Then
        pPaid = Format(pPaid, "yyyy-mm-dd") & "T" & Format(pPaid, "hh:mm:ss") & " +0800"
        strXML = strScope & pContNum & strChargeFor & pCharge & strPaid & pPaid & strGKey & pGKey & strSoapEnd
        Debug.Print strXML
        strOutput = SavePaidThruDay(strUrl, strSoapAction, strXML, strauthorization)
    End If

End Function

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
    SavePaidThruDay = "Error: " & err.Number & " - " & err.Description
MsgBox err.Description
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
    strOutput = ""
    strauthorization = strN4Authorization '"Basic bjRhcGk6d2VsY29tZQ=="
    strUrl = strN4Server & "/apex/services/argoservice?wsdl"
    strSoapAction = "POST " & "/apex/services/argoservice HTTP/1.1"
  
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
    objXmlHttp.Open "POST", strUrl, False, strN4UserName, strN4Password '"n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strauthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function

Public Function UpdateConsigneeAndBLNumber(ByVal pContNum As String, ByVal pConsignee As String, ByVal pBLNumber As String)
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
    strauthorization = strN4Authorization '"Basic bjRhcGk6d2VsY29tZQ=="
    strUrl = strN4Server & "/apex/services/argoservice?wsdl"
    'strSoapAction = "POST http://sbitc-apps1:9080/apex/services/argoservice HTTP/1.1"
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
"         <arg:xmlDoc><![CDATA[<icu><units><unit-identity id=""" & pContNum & """></unit-identity></units><properties>" & _
"<property tag=""GoodsConsignee"" value=""" & pConsignee & """/>" & _
"<property tag=""GoodsBlNbr"" value=""" & pBLNumber & """/>" & _
"</properties>" & _
"<event id=""UNIT_UPDATE_CONSIGNEE_API"" note=""ICU Update"" time-event-applied=""" & DateTime.Now & """ user-id=""n4api"" />" & _
"<event id=""UNIT_UPDATE_BILNUM_API"" note=""ICU Update"" time-event-applied=""" & DateTime.Now & """ user-id=""n4api"" />" & _
"</icu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, "n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strauthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function


Public Function AddWeighingHold(ByVal pContNum As String)
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
    strauthorization = "Basic bjRhcGk6d2VsY29tZQ=="
    'strUrl = "http://sbitc-apps1:9080/apex/services/argoservice?wsdl"
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
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit-identity id=""" & pContNum & """></unit-identity></units></entities><flags>" & _
"<flag hold-perm-id=""WEIGH_IMP_COMPLETE"" action=""ADD_HOLD""/>" & _
"</flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, "n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strauthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function

Public Function ReleaseOOGHold(ByVal pContNum As String)
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
    strauthorization = "Basic bjRhcGk6d2VsY29tZQ=="
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
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit-identity id=""" & pContNum & """></unit-identity></units></entities><flags>" & _
"<flag hold-perm-id=""BILLING_OOG"" action=""GRANT_PERMISSION""/>" & _
"</flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, "n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strauthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function


Public Function ReleaseDGHold(ByVal pContNum As String)
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
    strauthorization = "Basic bjRhcGk6d2VsY29tZQ=="
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
"         <arg:xmlDoc><![CDATA[<hpu><entities><units><unit-identity id=""" & pContNum & """></unit-identity></units></entities><flags>" & _
"<flag hold-perm-id=""BILLING_DG"" action=""GRANT_PERMISSION""/>" & _
"</flags></hpu>]]></arg:xmlDoc> " & _
"      </arg:genericInvoke> " & _
"   </soapenv:Body> " & _
"</soapenv:Envelope>"

    strXML = strScope
Debug.Print strXML

    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    

    ' Load XML
    objDom.async = False
    objDom.loadxml strXML
        
    ' Open the webservice
    objXmlHttp.Open "POST", strUrl, False, "n4api", "welcome"
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", strSoapAction
    objXmlHttp.setRequestHeader "Authorization", strauthorization
    
    ' Send XML command
    objXmlHttp.send objDom.xml
    strParam = objXmlHttp.responsetext
   Debug.Print strParam
    ' Close object
    Set objXmlHttp = Nothing
    
End Function




