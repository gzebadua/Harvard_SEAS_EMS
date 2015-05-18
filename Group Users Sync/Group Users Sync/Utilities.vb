Imports Microsoft.VisualBasic
Imports System.Web.Services
Imports System.Net.NetworkInformation
Imports System.Security.Cryptography
Imports System.Data
Imports System.Net.Mail
Imports System.Data.SqlClient


Public Module Utilities


    '========================================================================================
    'SQL Connections


    Dim servidor As String = "emsdbcluster.seas.harvard.edu"
    Dim servidorStaging As String = "emstestdb00.seas.harvard.edu"
    Dim usuario As String = "SEAS\ems-mc-ad-bind"
    Dim password As String = "4G6J@t!oPa"
    Dim database As String = "EMS_Staging"

    Public Function StringConnection() As String

        If Environ$("ComputerName").Contains("emstestfe00") = True Or Environ$("ComputerName").Contains("EMSTESTFE00") = True Then 'Staging servers

            Return "Data Source=" & servidorStaging & ";Initial Catalog=" & database & ";" & _
            "User Id=" & usuario & ";Password=" & password & ";Trusted_Connection=True;"

        Else 'Production

            Return "Data Source=" & servidor & ";Initial Catalog=" & database & ";" & _
            "User Id=" & usuario & ";Password=" & password & ";Trusted_Connection=True;"

        End If

    End Function


    '========================================================================================
    'Online Checks


    Public Function getSystemStatus() As String

        Dim output As String = ""

        If isSQLOnline() = False Then
            Return "SQL Offline"
        Else
            If isSystemOnline() = False Then
                Return "System Offline"
            Else
                Return "System Online"
            End If
        End If

    End Function


    Public Function isSQLOnline() As Boolean

        Dim result As String = ""

        result = getSQLQueryAsString("SELECT CURRENT_TIMESTAMP")

        If result.Equals("") Then
            Return False
        Else
            Return True
        End If

    End Function


    Public Function isSystemOnline() As Boolean

        Dim result As Integer = 0

        result = getSQLQueryAsInteger("SELECT CURRENT_TIMESTAMP")

        If result = 0 Then
            Return False
        Else
            Return True
        End If

    End Function


    '========================================================================================
    'SQL Helpers


    Public Function getSQLQueryAsDataset(ByVal query As String) As DataSet

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objDA As New SqlDataAdapter(query, objCon)
        Dim dsDatos As New DataSet

        Try

            objDA.Fill(dsDatos)

        Catch ex As Exception
            'Nothing
        Finally

            objDA.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

        Return dsDatos

    End Function


    Public Function getSQLQueryAsImage(ByVal query As String) As Byte()

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objDA As New SqlDataAdapter(query, objCon)
        Dim dsDatos As New DataSet

        Try

            objDA.Fill(dsDatos)

            Dim buffer As Byte() = DirectCast(dsDatos.Tables(0).Rows(0)("Data"), Byte())
            Return buffer

        Catch ex As Exception

            Return Nothing

        Finally

            objDA.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

    End Function


    Public Function get2SQLQueriesInSameDataset(ByVal query As String, ByVal query2 As String, ByVal nivel1PK As String, ByVal nivel2PK As String) As DataSet

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objDA As New SqlDataAdapter(query, objCon)
        Dim dsDatos As New DataSet

        Try

            objDA.Fill(dsDatos, "Level1")
            objDA = New SqlDataAdapter(query2, objCon)
            objDA.Fill(dsDatos, "Level2")
            dsDatos.Relations.Add("Children", dsDatos.Tables(0).Columns(nivel1PK), dsDatos.Tables(1).Columns(nivel2PK))

        Catch ex As Exception

            Dim exa As String
            exa = ex.ToString

        Finally

            objDA.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

        Return dsDatos

    End Function


    Public Function getSQLQueryAsString(ByVal query As String) As String

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objCmd As New SqlCommand(query, objCon)
        Dim objRdr As SqlDataReader

        Try

            objCon.Open()
            objRdr = objCmd.ExecuteReader
            objRdr.Read()
            If objRdr.HasRows Then
                Return objRdr(0)
            Else
                Return ""
            End If

        Catch ex As Exception

            Return ""

        Finally

            objCon.Close()
            objCon.Dispose()
            objCmd.Dispose()

        End Try

    End Function


    Public Function getSQLQueryAsInteger(ByVal query As String) As Integer

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objCmd As New SqlCommand(query, objCon)
        Dim objRdr As SqlDataReader

        Try

            objCon.Open()
            objRdr = objCmd.ExecuteReader
            objRdr.Read()
            If objRdr.HasRows Then
                Return CInt(objRdr(0))
            Else
                Return 0
            End If

        Catch ex As Exception

            Return 0

        Finally

            objCmd.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

    End Function


    Public Function getSQLQueryAsDouble(ByVal query As String) As Double

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objCmd As New SqlCommand(query, objCon)
        Dim objRdr As SqlDataReader

        Try

            objCon.Open()
            objRdr = objCmd.ExecuteReader
            objRdr.Read()
            If objRdr.HasRows Then
                Return CDbl(objRdr(0))
            Else
                Return 0.0
            End If

        Catch ex As Exception

            Return 0.0

        Finally

            objCmd.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

    End Function


    Public Function getSQLQueryAsBoolean(ByVal query As String) As Boolean

        Dim objCon As SqlConnection

        objCon = New SqlConnection(StringConnection())

        Dim objCmd As New SqlCommand(query, objCon)
        Dim objRdr As SqlDataReader

        Try

            objCon.Open()
            objRdr = objCmd.ExecuteReader
            objRdr.Read()
            If objRdr.HasRows Then
                Return CBool(objRdr(0))
            Else
                Return False
            End If

        Catch ex As Exception

            Return False

        Finally

            objCmd.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

    End Function


    Public Function executeSQLCommand(ByVal query As String) As Boolean

        Dim objCon As SqlConnection
        Dim objCmd As SqlCommand

        objCon = New SqlConnection(StringConnection())


        Try

            objCmd = New SqlCommand(query, objCon)

            objCon.Open()
            objCmd.CommandText = query
            objCmd.Connection = objCon
            objCmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception

            Return False

        Finally

            objCmd.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

    End Function


    Public Function executeTransactedSQLCommand(ByVal queries As String()) As Boolean

        Dim i As Integer = 0

        Dim objCon As SqlConnection
        Dim objCmd As SqlCommand

        objCon = New SqlConnection(StringConnection())

        Try

            objCmd = New SqlCommand("BEGIN TRANSACTION", objCon)
            objCon.Open()
            objCmd.CommandText = "BEGIN TRANSACTION"
            objCmd.Connection = objCon
            objCmd.CommandTimeout = 600
            objCmd.ExecuteNonQuery()

            For i = 0 To queries.Length - 1

                If queries(i) Is DBNull.Value Or queries(i) = "" Then
                    Continue For
                End If

                objCmd.CommandText = queries(i)
                objCmd.Connection = objCon
                objCmd.ExecuteNonQuery()

            Next i

            objCmd.CommandText = "COMMIT TRANSACTION"
            objCmd.Connection = objCon
            objCmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception

            'If ex.InnerException Is Nothing Then
            '    executeSQLCommand("INSERT IGNORE INTO errorlogs VALUES ('" & getMySQLDate() & "', 'The following query produced an exception: " & preventSQLInjection(queries(i)) & " : " & ex.ToString.Replace("'", "") & "')")
            'Else
            '    executeSQLCommand("INSERT IGNORE INTO errorlogs VALUES ('" & getMySQLDate() & "', 'The following query produced an exception: " & preventSQLInjection(queries(i)) & " : " & ex.ToString.Replace("'", "") & " Inner Exception: " & ex.InnerException.ToString.Replace("'", "") & "')")
            'End If

            'MsgBox(ex.ToString)

            Dim errorAtIteration As Integer = 0
            Dim verQueries As String = queries.Length
            errorAtIteration = i

            Try

                objCmd = New SqlCommand("ROLLBACK TRANSACTION", objCon)
                objCmd.CommandText = "ROLLBACK TRANSACTION"
                objCmd.Connection = objCon
                objCmd.ExecuteNonQuery()

            Catch ex2 As Exception

            End Try

            Return False

        Finally

            objCmd.Dispose()
            objCon.Close()
            objCon.Dispose()

        End Try

    End Function


    Public Function preventSQLInjection(ByVal value As String) As String

        If value.Contains("'") Then
            value = "'" & value & "'"
        End If

        Return value

    End Function


    '========================================================================================
    ' Time Functions


    Public Function getMSSQLDate() As String

        'Return getSQLQueryAsString("SELECT DATE_FORMAT(now(),'%Y%m%d')")
        Return getSQLQueryAsString("SELECT DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 3 HOUR),'%Y-%m-%d %H:%i:%s')")

    End Function


    Public Function getMSSQLTime() As String

        Return getSQLQueryAsString("SELECT DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 3 HOUR),'%H:%i:%s')")

    End Function


    '========================================================================================
    ' Security and Accessability Functions


    Private Function TryPing(ByVal Host As String) As Boolean

        Dim pingSender As New Ping
        Dim reply As PingReply

        Try

            reply = pingSender.Send(Host)

            If reply.Status = IPStatus.Success Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

            Return False

        End Try

    End Function


    Public Function findRemoteMachineName() As String

        Return Environ$("ComputerName")

    End Function


    '========================================================================================
    ' Mail Functions


    Public Function sendPlainMail(ByVal toEmail As String, ByVal subject As String, ByVal body As String) As Boolean

        Try

            Dim email As New MailMessage()
            Dim maFrom As New MailAddress("gzebadua@gmail.com", "SEAS Team")

            email.[To].Add(toEmail)
            email.From = maFrom
            email.Subject = subject
            email.Body = body


            Dim smtp As SmtpClient = New SmtpClient("smtp.gmail.com", 587) 'or 587 or 25
            smtp.Timeout = 30000
            smtp.Credentials = New System.Net.NetworkCredential(maFrom.Address, "AAAAAAAAA")

            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.EnableSsl = True

            smtp.Send(email)

            Return True

        Catch ex As Exception

            Try

                'SendMailAlt("smtp.gmail.com", 465, "gzebadua@gmail.com", "AAAAAA", "SEAS Team", "gzebadua@gmail.com", "", toEmail, subject, body, True)
                Return True
            Catch ex2 As Exception
                executeSQLCommand("INSERT INTO errorlogs VALUES ('" & getMSSQLDate() & "', '" & ex.ToString & "')")
                Return False
            End Try

        End Try

    End Function


    Public Function sendHTMLMail(ByVal toEmail As String, ByVal subject As String, ByVal body As String) As Boolean

        Try

            Dim email As New MailMessage()
            Dim maFrom As New MailAddress("gzebadua@gmail.com", "SEAS Team")

            email.[To].Add(toEmail)
            email.From = maFrom
            email.Subject = subject
            email.Body = body
            email.IsBodyHtml = True


            Dim smtp As SmtpClient = New SmtpClient("smtp.gmail.com", 587) 'or 587 or 25
            smtp.Timeout = 30000
            smtp.Credentials = New System.Net.NetworkCredential(maFrom.Address, "AAAAAAA")

            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.EnableSsl = True

            smtp.Send(email)

            Return True

        Catch ex As Exception

            Try

                'SendMailAlt("smtp.gmail.com", 465, "gzebadua@gmail.com", "AAAAA", "SEAS Team", "gzebadua@gmail.com", "", toEmail, subject, body, True)
                Return True
            Catch ex2 As Exception
                executeSQLCommand("INSERT INTO errorlogs VALUES ('" & getMSSQLDate() & "', '" & ex.ToString & "')")
                Return False
            End Try

        End Try

    End Function


End Module
