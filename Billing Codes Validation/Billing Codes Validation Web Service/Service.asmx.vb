Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Net
Imports System.IO
Imports System.Xml

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
<System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="SEAS")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service
    Inherits System.Web.Services.WebService


    Private Function validatefullBillingCodeStructure(ByVal fullBillingCode As String) As String()

        fullBillingCode = fullBillingCode.Trim

        Dim validation(8) As String

        validation(0) = "" 'Error
        validation(1) = "" 'Tub
        validation(2) = "" 'Org
        validation(3) = "" 'Obj
        validation(4) = "" 'Fund
        validation(5) = "" 'Activity
        validation(6) = "" 'SubActivity
        validation(7) = "" 'Root

        Dim separated As Boolean = False
        Dim separatedBillingCode As String()

        If fullBillingCode.Contains("-") = True Then
            separatedBillingCode = fullBillingCode.Split("-")
            separated = True
        End If

        If fullBillingCode.Contains(" ") = True Then
            separatedBillingCode = fullBillingCode.Split(" ")
            separated = True
        End If

        If separated = True Then

            For i = 0 To separatedBillingCode.Length - 1

                Select Case i

                    Case 0

                        If separatedBillingCode.GetValue(i).ToString.Length < 3 Or separatedBillingCode.GetValue(i).ToString.Length > 3 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "Tub should be 3 digits long"
                            End If
                        Else
                            validation(1) = separatedBillingCode.GetValue(i)
                        End If

                    Case 1

                        If separatedBillingCode.GetValue(i).ToString.Length < 5 Or separatedBillingCode.GetValue(i).ToString.Length > 5 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "Org should be 5 digits long"
                            End If
                        Else
                            validation(2) = separatedBillingCode.GetValue(i)
                        End If

                    Case 2

                        If separatedBillingCode.GetValue(i).ToString.Length < 4 Or separatedBillingCode.GetValue(i).ToString.Length > 4 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "Object should be 4 digits long"
                            End If
                        Else
                            validation(3) = separatedBillingCode.GetValue(i)
                        End If

                    Case 3

                        If separatedBillingCode.GetValue(i).ToString.Length < 6 Or separatedBillingCode.GetValue(i).ToString.Length > 6 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "Fund should be 6 digits long"
                            End If
                        Else
                            validation(4) = separatedBillingCode.GetValue(i)
                        End If

                    Case 4

                        If separatedBillingCode.GetValue(i).ToString.Length < 6 Or separatedBillingCode.GetValue(i).ToString.Length > 6 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "Activity should be 6 digits long"
                            End If
                        Else
                            validation(5) = separatedBillingCode.GetValue(i)
                        End If

                    Case 5

                        If separatedBillingCode.GetValue(i).ToString.Length < 4 Or separatedBillingCode.GetValue(i).ToString.Length > 4 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "SubActivity should be 4 digits long"
                            End If
                        Else
                            validation(6) = separatedBillingCode.GetValue(i)
                        End If

                    Case 6

                        If separatedBillingCode.GetValue(i).ToString.Length < 4 Or separatedBillingCode.GetValue(i).ToString.Length > 5 Or Char.IsNumber(separatedBillingCode.GetValue(i)) = False Then
                            If validation(0) = "" Then
                                validation(0) = "Root should be at least 4 digits long, 5 maximum"
                            End If
                        Else
                            validation(7) = separatedBillingCode.GetValue(i)
                        End If

                End Select

            Next i

        Else ' There are no separators

            Try
                validation(1) = fullBillingCode.Substring(0, 3)
            Catch ex As Exception
                If validation(0) = "" Then
                    validation(0) = "Tub should be 3 digits long"
                End If
            End Try

            Try
                validation(2) = fullBillingCode.Substring(3, 5)
            Catch ex As Exception
                If validation(0) = "" Then
                    validation(0) = "Org should be 5 digits long"
                End If
            End Try

            Try
                validation(3) = fullBillingCode.Substring(8, 4)
            Catch ex As Exception
                If validation(0) = "" Then
                    validation(0) = "Object should be 4 digits long"
                End If
            End Try

            Try
                validation(4) = fullBillingCode.Substring(12, 6)
            Catch ex As Exception
                If validation(0) = "" Then
                    validation(0) = "Fund should be 6 digits long"
                End If
            End Try

            Try
                validation(5) = fullBillingCode.Substring(18, 6)
            Catch ex As Exception
                If validation(0) = "" Then
                    validation(0) = "Activity should be 6 digits long"
                End If
            End Try

            Try
                validation(6) = fullBillingCode.Substring(24, 4)
            Catch ex As Exception
                If validation(0) = "" Then
                    validation(0) = "SubActivity should be 4 digits long"
                End If
            End Try

            Try
                validation(7) = fullBillingCode.Substring(28, 5)
            Catch ex As Exception

                Try
                    validation(7) = fullBillingCode.Substring(28, 4)
                Catch ex2 As Exception
                    If validation(0) = "" Then
                        validation(0) = "Root should be at least 4 digits long, 5 maximum"
                    End If
                End Try

            End Try

        End If

        separatedBillingCode = Nothing

        Return validation

    End Function


    <WebMethod(Description:="This is the method that our EMS installation (www.dea.com; http://events.seas.harvard.edu) uses to check for valid billing codes. This method calls validateBillingCodePartsReturnsStringAndErrors to validate the billing code and returns true or false depending on its validity. No errors codes or messages from the CoA validation service are returned with this function. It also calls validatefullBillingCodeStructure (internal private method) prior to sending sending easily verifiable incorrect data over the network to Central.")> _
    Public Function validateFullBillingCodeReturnsBoolean(ByVal billingCode As String) As Boolean

        Dim formatValidation() As String

        Dim response(1) As String

        response(0) = "N"
        response(1) = "Some error"

        formatValidation = validatefullBillingCodeStructure(billingCode)

        If formatValidation(0) = "" Then ' No Errors

            response = validateBillingCodePartsReturnsStringAndErrors(formatValidation(1), formatValidation(2), formatValidation(3), formatValidation(4), formatValidation(5), formatValidation(6), formatValidation(7))

            If response(0) = "Y" Then
                Return True
            Else
                Return False
            End If

        Else ' Errors found during validation

            Return False

        End If

    End Function


    <WebMethod(Description:="This method calls validateBillingCodePartsReturnsStringAndErrors to validate the billing code and returns true or false depending on its validity. This method returns the error messages from the CoA validation service (if any) so developers can relay them to the user. It also calls validatefullBillingCodeStructure (internal private method) prior to sending sending easily verifiable incorrect data over the network to Central.")> _
    Public Function validateFullBillingCodeReturnsStringAndErrors(ByVal billingCode As String) As String()

        Dim formatValidation() As String

        Dim response(1) As String

        response(0) = "N"
        response(1) = "Some error"

        formatValidation = validatefullBillingCodeStructure(billingCode)

        If formatValidation(0) = "" Then ' No Errors

            response = validateBillingCodePartsReturnsStringAndErrors(formatValidation(1), formatValidation(2), formatValidation(3), formatValidation(4), formatValidation(5), formatValidation(6), formatValidation(7))


        Else ' Errors found during validation

            response(0) = "N"
            response(1) = formatValidation(0)

        End If

        Return response

    End Function


    <WebMethod(Description:="This method calls validateBillingCodeReturnsStringAndErrors to validate the billing code and returns true or false depending on its validity. No errors codes or messages from the CoA validation service are returned with this function. It also calls validatefullBillingCodeStructure (internal private method) prior to sending sending easily verifiable incorrect data over the network to Central.")> _
    Public Function validateBillingCodePartsReturnsBoolean(ByVal tub As String, ByVal org As String, ByVal obj As String, ByVal fund As String, ByVal activity As String, ByVal subactivity As String, ByVal root As String) As Boolean

        Dim formatValidation() As String

        Dim response(1) As String

        response(0) = "N"
        response(1) = "Some error"

        formatValidation = validatefullBillingCodeStructure(tub & org & obj & fund & activity & subactivity & root)

        If formatValidation(0) = "" Then ' No Errors

            'response = validateBillingCodePartsReturnsStringAndErrors(tub, org, obj, fund, activity, subactivity, root) 'Same thing, I just like reusing the functions...
            response = validateBillingCodePartsReturnsStringAndErrors(formatValidation(1), formatValidation(2), formatValidation(3), formatValidation(4), formatValidation(5), formatValidation(6), formatValidation(7))

            If response(0) = "Y" Then
                Return True
            Else
                Return False
            End If

        Else ' Errors found during validation

            Return False

        End If

    End Function


    <WebMethod(Description:="This method validates the billing code provided against the CoA validation service and returns the error messages (if any) so developers can relay them to the user. It also calls validatefullBillingCodeStructure (internal private method) prior to sending sending easily verifiable incorrect data over the network to Central.")> _
    Public Function validateBillingCodePartsReturnsStringAndErrors(ByVal tub As String, ByVal org As String, ByVal obj As String, ByVal fund As String, ByVal activity As String, ByVal subactivity As String, ByVal root As String) As String()

        Dim formatValidation() As String

        Dim response(1) As String

        response(0) = "N"
        response(1) = "Some error"

        formatValidation = validatefullBillingCodeStructure(tub & org & obj & fund & activity & subactivity & root)

        If formatValidation(0) = "" Then ' No Errors

            Dim CoA_URL As String = "https://apollo36.cadm.harvard.edu:8052/GLValidate/ValidateGLAccount"

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

            Using soapResponse As WebResponse = req.GetResponse()

                Using rd As New StreamReader(soapResponse.GetResponseStream())

                    Dim soapResult As String = rd.ReadToEnd()
                    response(0) = soapResult.Substring(soapResult.IndexOf("<ValidationFlag>") + 16, 1)
                    response(1) = soapResult.Substring(soapResult.IndexOf("<ErrorMessage>") + 14, soapResult.IndexOf("</ErrorMessage>") - (soapResult.IndexOf("<ErrorMessage>") + 14))

                End Using

            End Using

            req = Nothing
            CoA_URL = Nothing

        Else ' Errors found during validation

            response(0) = "N"
            response(1) = formatValidation(0)

        End If

        Return response

    End Function


End Class