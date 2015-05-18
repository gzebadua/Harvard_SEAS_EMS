Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Net
Imports System.IO
Imports System.Xml

Namespace SEAS

    ' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
    <System.Web.Script.Services.ScriptService()> _
    <System.Web.Services.WebService(Namespace:="SEAS")> _
    <System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
    <ToolboxItem(False)> _
    Public Class BillingCodes
        Inherits System.Web.Services.WebService


        <WebMethod()> _
        Public Function validateBillingCode(ByVal tub As String, ByVal org As String, ByVal obj As String, ByVal fund As String, ByVal activity As String, ByVal subactivity As String, ByVal root As String) As Boolean


            Dim CoA_URL As String = "https://apollo36.cadm.harvard.edu:8052/GLValidate/ValidateGLAccount"

            Dim validationFlag As String = "N"

            Dim isBillingCodeValid As Boolean = False

            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(CoA_URL), HttpWebRequest)
            req.Headers.Add("SOAP:Action")
            req.ContentType = "text/xml;charset=""utf-8"""
            req.Accept = "text/xml"
            req.Method = "POST"

            Dim soapEnvelopeXml As New XmlDocument()
            soapEnvelopeXml.LoadXml("<?xml version=""1.0"" encoding=""utf-8""?>" & _
                                    "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" & _
                                    "<soap:Body>" & _
                                    "    <ValidateGLAccountRequest>" & _
                                    "        <Tub>" & tub & "</Tub>" & _
                                    "        <Org>" & org & "</Org>" & _
                                    "        <Object>" & obj & "</Object>" & _
                                    "        <Fund>" & fund & "</Fund>" & _
                                    "        <Activity>" & activity & "</Activity>" & _
                                    "        <Subactivity>" & subactivity & "</Subactivity>" & _
                                    "        <Root>" & root & "</Root>" & _
                                    "    </ValidateGLAccountRequest>" & _
                                    "</soap:Body>" & _
                                    "</soap:Envelope>")

            Using stream As Stream = req.GetRequestStream()
                soapEnvelopeXml.Save(stream)
            End Using

            Using response As WebResponse = req.GetResponse()

                Using rd As New StreamReader(response.GetResponseStream())

                    Dim soapResult As String = rd.ReadToEnd()
                    validationFlag = soapResult.Substring(soapResult.IndexOf("<ValidationFlag>") + 16, 1)

                    If validationFlag = "Y" Then
                        isBillingCodeValid = True
                    Else
                        isBillingCodeValid = False
                    End If

                End Using

            End Using

            req = Nothing
            CoA_URL = Nothing
            validationFlag = Nothing

            Return isBillingCodeValid

        End Function


        <WebMethod()> _
        Public Function customBillingReferenceValidation(ByVal billingCode As String) As Boolean  ' This is the method that EMS uses to check for valid billing codes. We are parsing the string and then using the validateBillingCode webmethod to keep everything centralized and easy to modify if necessary

            billingCode = billingCode.Trim

            If billingCode.Contains("-") = True Or billingCode.Contains(" ") = True Then
                billingCode = billingCode.Replace("-", "").Replace(" ", "")
            End If

            If billingCode.Length < 33 Then
                Return False
            End If

            Dim tub As String = ""
            Dim org As String = ""
            Dim obj As String = ""
            Dim fund As String = ""
            Dim activity As String = ""
            Dim subactivity As String = ""
            Dim root As String = ""

            Try

                tub = billingCode.Substring(0, 3)
                org = billingCode.Substring(3, 5)
                obj = billingCode.Substring(8, 4)
                fund = billingCode.Substring(12, 6)
                activity = billingCode.Substring(18, 6)
                subactivity = billingCode.Substring(24, 4)
                root = billingCode.Substring(28, 5)

            Catch ex As Exception

                Return False 'If something goes wrong here it means that the input data is wrong, therefore billingCode invalid

            End Try

            Dim isBillingCodeValid As Boolean = False

            isBillingCodeValid = validateBillingCode(tub, org, obj, fund, activity, subactivity, root)

            tub = Nothing
            org = Nothing
            obj = Nothing
            fund = Nothing
            activity = Nothing
            subactivity = Nothing
            root = Nothing

            Return isBillingCodeValid

        End Function


        <WebMethod()> _
        Public Function customPONumberValidation(ByVal billingCode As String) As Boolean

            Return True

        End Function


    End Class

End Namespace