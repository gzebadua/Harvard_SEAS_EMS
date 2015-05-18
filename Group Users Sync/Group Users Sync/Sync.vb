Imports System.DirectoryServices


Public Class Sync


    Private Sub Sync_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        Application.DoEvents()
        SyncADUserInfoToEMSUsers()
        System.Environment.Exit(0)

    End Sub


    Private Sub Sync_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        

    End Sub


    Public Function getAllUsers() As DataSet

        Dim ADPath As String = "LDAP://dc5.seas.harvard.edu/ou=people,dc=seas,dc=harvard,dc=edu"
        Dim filter As String = "(&(objectClass=user)(objectCategory=person))"

        Dim ds As New DirectorySearcher

        Dim dsPeople As New DataSet()
        Dim tbPeople As New DataTable()

        dsPeople.Tables.Add(tbPeople)

        ds.SearchRoot = New DirectoryEntry(ADPath, "SEAS\ems-mc-ad-bind", "4G6J@t!oPa", DirectoryServices.AuthenticationTypes.Secure)
        ds.SearchScope = SearchScope.Subtree

        ds.PropertiesToLoad.Add("sAMAccountName")
        ds.PropertiesToLoad.Add("givenName")
        ds.PropertiesToLoad.Add("sn")
        ds.PropertiesToLoad.Add("title")
        ds.PropertiesToLoad.Add("harvardEduMiddlename")
        ds.PropertiesToLoad.Add("mail")
        ds.PropertiesToLoad.Add("telephoneNumber")
        ds.PropertiesToLoad.Add("facsimileTelephoneNumber")
        ds.PropertiesToLoad.Add("streetAddress")
        ds.PropertiesToLoad.Add("buildingName")
        ds.PropertiesToLoad.Add("l")
        ds.PropertiesToLoad.Add("st")
        ds.PropertiesToLoad.Add("postalCode")
        ds.PropertiesToLoad.Add("c")
        ds.PropertiesToLoad.Add("sAMAccountName")
        ds.PropertiesToLoad.Add("harvardEduADHUID")
        'ds.PropertiesToLoad.Add("memberOf")                 'Commented out after DEA modified the HRTK to disable group sync, on our request
        'ds.PropertiesToLoad.Add("eduPersonAffiliation")

        tbPeople.Columns.Add("employeeid", GetType(String))
        tbPeople.Columns.Add("firstName", GetType(String))
        tbPeople.Columns.Add("lastName", GetType(String))
        tbPeople.Columns.Add("title", GetType(String))
        tbPeople.Columns.Add("middleName", GetType(String))
        tbPeople.Columns.Add("mail", GetType(String))
        tbPeople.Columns.Add("telephoneNumber", GetType(String))
        tbPeople.Columns.Add("facsimileTelephoneNumber", GetType(String))
        tbPeople.Columns.Add("streetAddress", GetType(String))
        tbPeople.Columns.Add("streetAddress2", GetType(String))
        tbPeople.Columns.Add("city", GetType(String))
        tbPeople.Columns.Add("state", GetType(String))
        tbPeople.Columns.Add("zipcode", GetType(String))
        tbPeople.Columns.Add("country", GetType(String))
        'tbPeople.Columns.Add("billingReference", GetType(String))                 'Added by DEA, Leave it NULL  'Commented after Nemo HRTK
        tbPeople.Columns.Add("networkReference", GetType(String))
        'tbPeople.Columns.Add("groupid", GetType(String))                 'Commented out after DEA modified the HRTK to disable group sync, on our request
        'tbPeople.Columns.Add("groupName", GetType(String))
        'tbPeople.Columns.Add("groupType", GetType(String))
        tbPeople.Columns.Add("huid", GetType(String))

        ds.Filter = filter

        For Each sr As SearchResult In ds.FindAll()

            Dim row As DataRow = tbPeople.NewRow()

            Try
                row("employeeid") = sr.Properties("sAMAccountName").Item(0).ToString.Substring(0, 19)
            Catch ex As Exception

                Try
                    row("employeeid") = sr.Properties("sAMAccountName").Item(0)
                Catch ex2 As Exception
                    row("employeeid") = ""
                End Try

            End Try

            Try
                row("firstName") = sr.Properties("givenName").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("firstName") = sr.Properties("givenName").Item(0)
                Catch ex2 As Exception
                    row("firstName") = ""
                End Try

            End Try

            Try
                row("lastName") = sr.Properties("sn").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("lastName") = sr.Properties("sn").Item(0)
                Catch ex2 As Exception
                    row("lastName") = ""
                End Try

            End Try

            Try
                row("title") = sr.Properties("title").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("title") = sr.Properties("title").Item(0)
                Catch ex2 As Exception
                    row("title") = ""
                End Try

            End Try

            Try
                row("middleName") = sr.Properties("harvardEduMiddlename").Item(0).ToString.Substring(0, 9)
            Catch ex As Exception

                Try
                    row("middleName") = sr.Properties("harvardEduMiddlename").Item(0)
                Catch ex2 As Exception
                    row("middleName") = ""
                End Try

            End Try

            Try
                row("mail") = sr.Properties("mail").Item(0).ToString.Substring(0, 74)
            Catch ex As Exception

                Try
                    row("mail") = sr.Properties("mail").Item(0)
                Catch ex2 As Exception
                    row("mail") = ""
                End Try

            End Try

            Try
                row("telephoneNumber") = sr.Properties("telephoneNumber").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("telephoneNumber") = sr.Properties("telephoneNumber").Item(0)
                Catch ex2 As Exception
                    row("telephoneNumber") = ""
                End Try

            End Try

            Try
                row("facsimileTelephoneNumber") = sr.Properties("facsimileTelephoneNumber").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("facsimileTelephoneNumber") = sr.Properties("facsimileTelephoneNumber").Item(0)
                Catch ex2 As Exception
                    row("facsimileTelephoneNumber") = ""
                End Try

            End Try

            Try
                row("streetAddress") = sr.Properties("streetAddress").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("streetAddress") = sr.Properties("streetAddress").Item(0)
                Catch ex2 As Exception
                    row("streetAddress") = ""
                End Try

            End Try

            Try
                row("streetAddress2") = sr.Properties("buildingName").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("streetAddress2") = sr.Properties("buildingName").Item(0)
                Catch ex2 As Exception
                    row("streetAddress2") = ""
                End Try

            End Try

            Try
                row("city") = sr.Properties("l").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("city") = sr.Properties("l").Item(0)
                Catch ex2 As Exception
                    row("city") = ""
                End Try

            End Try

            Try
                row("state") = sr.Properties("st").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("state") = sr.Properties("st").Item(0)
                Catch ex2 As Exception
                    row("state") = ""
                End Try

            End Try

            Try
                row("zipcode") = sr.Properties("postalCode").Item(0).ToString.Substring(0, 9)
            Catch ex As Exception

                Try
                    row("zipcode") = sr.Properties("postalCode").Item(0)
                Catch ex2 As Exception
                    row("zipcode") = ""
                End Try

            End Try

            Try
                row("country") = sr.Properties("c").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("country") = sr.Properties("c").Item(0)
                Catch ex2 As Exception
                    row("country") = ""
                End Try

            End Try

            'Try                                              'Added by DEA, Leave it NULL  'Commented after Nemo HRTK
            '    row("billingReference") = DBNull.Value
            'Catch ex As Exception
            '    row("billingReference") = ""
            'End Try


            Try
                row("networkReference") = ("SEAS\" & sr.Properties("sAMAccountName").Item(0)).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("networkReference") = "SEAS\" & sr.Properties("sAMAccountName").Item(0)
                Catch ex2 As Exception
                    row("networkReference") = ""
                End Try

            End Try

            'Dim tmpGroups As String = ""                 'Commented out after DEA modified the HRTK to disable group sync, on our request

            'Try

            '    If sr.Properties("memberOf").Count = 1 Then 'Users with only 1 group, e.g. their own group (same samaccountname for the group: gzebadua gzebadua)
            '        tmpGroups = tmpGroups & sr.Properties("memberOf").Item(0)
            '        Exit Try
            '    End If

            '    For i = 0 To sr.Properties("memberOf").Count - 1 'Members with more than 1 group, save the one that is not the same as the samaccountname of the member.

            '        If sr.Properties("memberOf").Item(i) <> sr.Properties("sAMAccountName").Item(0) Then
            '            tmpGroups = tmpGroups & sr.Properties("memberOf").Item(i)
            '            Exit Try
            '        End If

            '    Next i

            'Catch ex As Exception

            'End Try

            'tmpGroups = tmpGroups.Replace("CN=", "").Replace("DC=seas,DC=harvard,DC=edu", "").Replace("OU=groups,", "")
            'tmpGroups = tmpGroups.Replace("OU=ad-local,", "").Replace("OU=admin-group-import,", "").Replace("OU=committees,", "").Replace("OU=departments,", "").Replace("OU=research,", "").Replace("OU=test,", "").Replace("OU=unix,", "")
            'tmpGroups = tmpGroups.Replace("OU=vmware-admin-delegation-group,", "").Replace("OU=papercut,", "").Replace("OU=SEAS-Exch-MailGroups,", "").Replace("OU=Test-mig,", "")

            'If tmpGroups.EndsWith(",") Then 'To remove the trailing comma at the end
            '    tmpGroups = tmpGroups.Substring(0, tmpGroups.Length - 1)
            'End If

            'If tmpGroups = "" Then
            '    tmpGroups = "noneSet"
            'End If

            'row("groupid") = tmpGroups

            'Try

            '    If tmpGroups = "noneSet" Then
            '        row("groupName") = "No Group Name Set Yet"
            '    Else
            '        row("groupName") = GetGroupName(tmpGroups).Substring(0, 49)
            '    End If

            'Catch ex As Exception

            '    Try
            '        row("groupName") = GetGroupName(tmpGroups)
            '    Catch ex2 As Exception
            '        row("groupName") = "No Group Name Set Yet"
            '    End Try

            'End Try

            'Dim tmpGroupType As String = ""

            'If tmpGroups = "noneSet" Then

            '    tmpGroupType = "Staff"

            'Else

            '    Try
            '        tmpGroupType = sr.Properties("eduPersonAffiliation").Item(0)
            '    Catch ex As Exception
            '        tmpGroupType = "Staff"
            '    End Try

            '    If tmpGroupType = "Graduate Student" Or tmpGroupType = "Undergraduate Intern" Or tmpGroupType = "Visiting Undergraduate Intern" Then
            '        tmpGroupType = "Student"
            '    End If

            '    If tmpGroupType = "Tenured Professor" Or tmpGroupType = "Emeritus" Then
            '        tmpGroupType = "Faculty"
            '    End If

            '    If tmpGroupType = "Associate" Or tmpGroupType = "Associate of SEAS" Then
            '        tmpGroupType = "Affiliate"
            '    End If

            '    If tmpGroupType = "Senior Research Fellow" Or tmpGroupType = "Research Associate" Or tmpGroupType = "Visiting Undergraduate Research Intern" Then
            '        tmpGroupType = "Researcher"
            '    End If

            '    If tmpGroupType = "Consultant" Then
            '        tmpGroupType = "External"
            '    End If

            '    If tmpGroupType = "Exempt Staff" Or tmpGroupType = "Non-Exempt Staff" Or tmpGroupType = "Visiting Faculty" Or tmpGroupType = "Visiting Professor" Or tmpGroupType = "Visiting Scholar" Or tmpGroupType = "Preceptor" Or tmpGroupType = "Fellow" Or tmpGroupType = "Postdoctoral Fellow" Or tmpGroupType = "Senior Lecturer" Or tmpGroupType = "Visiting Lecturer" Or tmpGroupType = "Lecturer" Or tmpGroupType = "Senior Preceptor" Or tmpGroupType = "Assistant Professor" Or tmpGroupType = "Associate Professor" Or tmpGroupType = "Professor of the Practice" Then
            '        tmpGroupType = "Staff"
            '    End If

            '    If tmpGroupType = "" Then
            '        tmpGroupType = "Staff"
            '    End If

            'End If

            'Try

            '    row("groupType") = tmpGroupType.Substring(0, 29)

            'Catch ex As Exception

            '    row("groupType") = tmpGroupType

            'End Try


            Try
                row("huid") = sr.Properties("harvardEduADHUID").Item(0).ToString.Substring(0, 49)
            Catch ex As Exception

                Try
                    row("huid") = sr.Properties("harvardEduADHUID").Item(0)
                Catch ex2 As Exception
                    row("huid") = ""
                End Try

            End Try

            tbPeople.Rows.Add(row)

        Next

        ds.Dispose()

        Return dsPeople


    End Function


    Public Function getAllGroups() As DataSet

        Dim ADPath As String = "LDAP://dc5.seas.harvard.edu/ou=groups,dc=seas,dc=harvard,dc=edu"
        Dim filter As String = "(objectCategory=group)"

        Dim ds As New DirectorySearcher
        Dim results As New List(Of DirectoryEntry)

        Dim dsGroups As New DataSet()
        Dim tbGroups As New DataTable()

        dsGroups.Tables.Add(tbGroups)

        ds.SearchRoot = New DirectoryEntry(ADPath, "SEAS\admgzebadua", "DLF1594g!", DirectoryServices.AuthenticationTypes.Secure)
        ds.SearchScope = SearchScope.Subtree

        ds.PropertiesToLoad.Add("sAMAccountName")
        ds.PropertiesToLoad.Add("description")

        tbGroups.Columns.Add("groupid", GetType(String))
        tbGroups.Columns.Add("groupName", GetType(String))

        ds.Filter = filter

        For Each sr As SearchResult In ds.FindAll()

            Dim row As DataRow = tbGroups.NewRow()

            Try
                row("groupid") = sr.Properties("sAMAccountName").Item(0)
            Catch ex As Exception
                row("groupid") = ""
            End Try

            Try
                row("groupName") = sr.Properties("description").Item(0)
            Catch ex As Exception

                Try
                    row("groupName") = sr.Properties("sAMAccountName").Item(0)
                Catch ex2 As Exception
                    row("groupName") = ""
                End Try

            End Try

            tbGroups.Rows.Add(row)

        Next

        ds.Dispose()

        Return dsGroups


    End Function


    Public Function SyncUsers() As Boolean

        Dim dsSyncData As DataSet = getAllUsers()

        Dim queries(dsSyncData.Tables(0).Rows.Count) As String
        Dim queriesHUID(dsSyncData.Tables(0).Rows.Count) As String

        Dim strPeopleTableColumnsQuery As String

        strPeopleTableColumnsQuery = "" & _
        "SELECT column_name = syscolumns.name " & _
        "FROM dbo.sysobjects " & _
        "JOIN dbo.syscolumns ON sysobjects.id = syscolumns.id " & _
        "JOIN dbo.systypes ON syscolumns.xusertype=systypes.xusertype " & _
        "WHERE sysobjects.xtype='U' " & _
        "AND sysobjects.name = 'tblPeople' " & _
        "ORDER BY sysobjects.name,syscolumns.colid "

        Dim dsPeopleTableColumns As DataSet = getSQLQueryAsDataset(strPeopleTableColumnsQuery)

        For i = 0 To dsSyncData.Tables(0).Rows.Count - 1

            'If user already exists in EMS_Staging.dbo.tblPeople, Update, else Insert. DO NOT TRUNCATE TABLE!!!!!

            If getSQLQueryAsString("SELECT * FROM EMS_Staging.dbo.tblPeople WHERE PersonnelNumber = '" & dsSyncData.Tables(0).Rows(i).Item(0).ToString.Replace("'", "").Replace("--", "") & "'") <> "" Then

                queries(i) = "UPDATE EMS_Staging.dbo.tblPeople SET "

                For j = 0 To dsSyncData.Tables(0).Columns.Count - 2 'It used to be -1 but we added the HUID column

                    If j = dsSyncData.Tables(0).Columns.Count - 2 Then 'Last column. Note: see comment above
                        queries(i) = queries(i) & dsPeopleTableColumns.Tables(0).Rows(j).Item(0).ToString & " = '" & dsSyncData.Tables(0).Rows(i).Item(j).ToString.Replace("'", "").Replace("--", "") & "' "
                    Else
                        queries(i) = queries(i) & dsPeopleTableColumns.Tables(0).Rows(j).Item(0).ToString & " = '" & dsSyncData.Tables(0).Rows(i).Item(j).ToString.Replace("'", "").Replace("--", "") & "', "
                    End If

                Next j

                queries(i) = queries(i) & "WHERE PersonnelNumber = '" & dsSyncData.Tables(0).Rows(i).Item(0).ToString.Replace("'", "").Replace("--", "") & "'"

            Else

                queries(i) = "INSERT INTO EMS_Staging.dbo.tblPeople VALUES ('"

                For j = 0 To dsSyncData.Tables(0).Columns.Count - 2 'It used to be -1 but we added the HUID column

                    If j = dsSyncData.Tables(0).Columns.Count - 2 Then 'Last column. Note: see comment above
                        queries(i) = queries(i) & dsSyncData.Tables(0).Rows(i).Item(j).ToString.Replace("'", "").Replace("--", "") & "')"
                    Else
                        queries(i) = queries(i) & dsSyncData.Tables(0).Rows(i).Item(j).ToString.Replace("'", "").Replace("--", "") & "', '"
                    End If

                Next j

            End If


        Next i

        For j = 0 To dsSyncData.Tables(0).Rows.Count - 1


            If getSQLQueryAsString("SELECT * FROM EMS_Staging.dbo.tblHUIDs WHERE AD = '" & dsSyncData.Tables(0).Rows(j).Item(0).ToString.Replace("'", "").Replace("--", "") & "' AND HUID = ''") <> "" Then

                queriesHUID(j) = "UPDATE EMS_Staging.dbo.tblHUIDs SET HUID = " & dsSyncData.Tables(0).Rows(j).Item(dsSyncData.Tables(0).Columns.Count - 1).ToString.Replace("'", "").Replace("--", "") & " WHERE AD = '" & dsSyncData.Tables(0).Rows(j).Item(0).ToString.Replace("'", "").Replace("--", "") & "'"

            Else

                queriesHUID(j) = "INSERT INTO EMS_Staging.dbo.tblHUIDs VALUES ('" & dsSyncData.Tables(0).Rows(j).Item(0).ToString.Replace("'", "").Replace("--", "") & "', '" & dsSyncData.Tables(0).Rows(j).Item(dsSyncData.Tables(0).Columns.Count - 1).ToString.Replace("'", "").Replace("--", "") & "')"
                
            End If


        Next j

        If executeTransactedSQLCommand(queries) = True Then

            executeTransactedSQLCommand(queriesHUID)

            Return True

        Else

            Return False

        End If

    End Function


    Public Function DeleteInactiveUsers() As Boolean

        Dim dsSyncData As DataSet = getDisabledUsers()

        Dim tmpRowCount As Integer = dsSyncData.Tables(0).Rows.Count
        Dim queries(tmpRowCount) As String

        For i = 0 To dsSyncData.Tables(0).Rows.Count - 1

            queries(i) = "DELETE FROM EMS_Staging.dbo.tblPeople WHERE PersonnelNumber = '" & dsSyncData.Tables(0).Rows(i).Item(0).ToString.Replace("'", "").Replace("--", "") & "'"

        Next i

        If executeTransactedSQLCommand(queries) = True Then

            Return True

        Else

            Return False

        End If

    End Function


    Public Function getDisabledUsers() As DataSet

        Dim ADPath As String = "LDAP://dc5.seas.harvard.edu/ou=people,dc=seas,dc=harvard,dc=edu"
        Dim filter As String = "(&(objectClass=user)(objectCategory=person)(userAccountControl=514))"

        Dim dsToDelete As New DirectorySearcher

        Dim dsPeopleToDelete As New DataSet()
        Dim tbPeopleToDelete As New DataTable()

        dsPeopleToDelete.Tables.Add(tbPeopleToDelete)

        dsToDelete.SearchRoot = New DirectoryEntry(ADPath, "SEAS\ems-mc-ad-bind", "4G6J@t!oPa", DirectoryServices.AuthenticationTypes.Secure)
        dsToDelete.SearchScope = SearchScope.Subtree

        dsToDelete.PropertiesToLoad.Add("sAMAccountName")

        tbPeopleToDelete.Columns.Add("employeeid", GetType(String))

        dsToDelete.Filter = filter

        For Each sr As SearchResult In dsToDelete.FindAll()

            Dim row As DataRow = tbPeopleToDelete.NewRow()

            Try
                row("employeeid") = sr.Properties("sAMAccountName").Item(0).ToString.Substring(0, 19)
            Catch ex As Exception

                Try
                    row("employeeid") = sr.Properties("sAMAccountName").Item(0)
                Catch ex2 As Exception
                    row("employeeid") = ""
                End Try

            End Try

            tbPeopleToDelete.Rows.Add(row)

        Next

        dsToDelete.Dispose()

        Return dsPeopleToDelete


    End Function


    Public Function DeleteUnwantedUsers() As Boolean

        Dim unwantedQuery As String = "SELECT PersonnelNumber FROM EMS_Staging.dbo.tblPeople WHERE " & _
        "PersonnelNumber = 'science_cooking' OR " & _
        "PersonnelNumber = 'cscie207' OR " & _
        "PersonnelNumber = 'cs121' OR " & _
        "PersonnelNumber = 'cs50' OR " & _
        "PersonnelNumber = 'cs246' OR " & _
        "PersonnelNumber = 'cs221' OR " & _
        "PersonnelNumber = 'itacctb' OR " & _
        "PersonnelNumber = 'admissions' OR " & _
        "PersonnelNumber = 'amam107' OR " & _
        "PersonnelNumber = 'crcs-apply' OR " & _
        "PersonnelNumber = 'search' OR " & _
        "PersonnelNumber = 'classes' OR " & _
        "PersonnelNumber = 'roomscopy' OR " & _
        "PersonnelNumber = 'jira-help' OR " & _
        "PersonnelNumber = 'www-ho' OR " & _
        "PersonnelNumber = 'physics95_mazur' OR " & _
        "PersonnelNumber = 'micro-web' OR " & _
        "PersonnelNumber = 'umc-chip' OR " & _
        "PersonnelNumber = 'nsec-web' OR " & _
        "PersonnelNumber = 'itpgp' OR " & _
        "PersonnelNumber = 'researchpositions' OR " & _
        "PersonnelNumber = 'arc-purchasing' OR " & _
        "PersonnelNumber = 'rooms' OR " & _
        "PersonnelNumber = 'energy-search' OR " & _
        "PersonnelNumber = 'softmat-web' OR " & _
        "PersonnelNumber = 'teachinglabs-admins' OR " & _
        "PersonnelNumber = 'plonetest' OR " & _
        "PersonnelNumber = 'jftest25' OR " & _
        "PersonnelNumber = 'jftest26' OR " & _
        "PersonnelNumber = 'jftest27' OR " & _
        "PersonnelNumber = 'jftest28' OR " & _
        "PersonnelNumber = 'softmatter_tf' OR " & _
        "PersonnelNumber = 'it-training' OR " & _
        "PersonnelNumber = 'fasfc-corp-outreach' OR " & _
        "PersonnelNumber = 'financeuser' OR " & _
        "PersonnelNumber = 'monitor' OR " & _
        "PersonnelNumber = 'physics95_dvw' OR " & _
        "PersonnelNumber = 'softmatter_dvw' OR " & _
        "PersonnelNumber = 'plone-Photonics' OR " & _
        "PersonnelNumber = 'liquids-web' OR " & _
        "PersonnelNumber = 'ADMIN-itps' OR " & _
        "PersonnelNumber = 'ADMIN-iic-group' OR " & _
        "PersonnelNumber = 'iis' OR " & _
        "PersonnelNumber = 'fs-hips' OR " & _
        "PersonnelNumber = 'vision' OR " & _
        "PersonnelNumber = 'ADMIN-FinShare Group' OR " & _
        "PersonnelNumber = 'ADMIN-FinShare' OR " & _
        "PersonnelNumber = 'hugroup' OR " & _
        "PersonnelNumber = 'ADMIN-entech' OR " & _
        "PersonnelNumber = 'envmicro-web' OR " & _
        "PersonnelNumber = 'jira' OR " & _
        "PersonnelNumber = 'jira-internal' OR " & _
        "PersonnelNumber = 'eguideweb' OR " & _
        "PersonnelNumber = 'ADMIN-auguste_group' OR " & _
        "PersonnelNumber = 'ADMIN-weitz-group' OR " & _
        "PersonnelNumber = 'culturomics' OR " & _
        "PersonnelNumber = 'joint_council' OR " & _
        "PersonnelNumber = 'mitchell02' OR " & _
        "PersonnelNumber = 'climate-web' OR " & _
        "PersonnelNumber = 'ADMIN-BMR-Users' OR " & _
        "PersonnelNumber = 'ADMIN-BASF' OR " & _
        "PersonnelNumber = 'ADMIN-admissions' OR " & _
        "PersonnelNumber = 'ADMIN-itl' OR " & _
        "PersonnelNumber = 'ADMIN-martin-group' OR " & _
        "PersonnelNumber = 'ADMIN-hr-payroll' OR " & _
        "PersonnelNumber = 'ADMIN-sdr' OR " & _
        "PersonnelNumber = 'ADMIN-suo-group' OR " & _
        "PersonnelNumber = 'ADMIN-vlassak' OR " & _
        "PersonnelNumber = 'ADMIN-WYSS-share-users' OR " & _
        "PersonnelNumber = 'admsgrp' OR " & _
        "PersonnelNumber = 'softmatter-manohara' OR " & _
        "PersonnelNumber = 'plone-ChinaProject' OR " & _
        "PersonnelNumber = 'plone-GSLC' OR " & _
        "PersonnelNumber = 'plone-Hemo' OR " & _
        "PersonnelNumber = 'plone-JoshiGroup' OR " & _
        "PersonnelNumber = 'plone-nsfgrant' OR " & _
        "PersonnelNumber = 'academicoffice' OR " & _
        "PersonnelNumber = 'communicationsoffic' OR " & _
        "PersonnelNumber = 'cs-ad' OR " & _
        "PersonnelNumber = 'costsavings' OR " & _
        "PersonnelNumber = 'es139' OR " & _
        "PersonnelNumber = 'gridwebteam' OR " & _
        "PersonnelNumber = 'infosec' OR " & _
        "PersonnelNumber = 'networks' OR " & _
        "PersonnelNumber = 'pirp' OR " & _
        "PersonnelNumber = 'venky-assistant' OR " & _
        "PersonnelNumber = 'murray-assistant' OR " & _
        "PersonnelNumber = 'camurray-test' OR " & _
        "PersonnelNumber = 'mfreeman-test' OR " & _
        "PersonnelNumber = 'ParkingDeansOffice' OR " & _
        "PersonnelNumber = 'nectar' OR " & _
        "PersonnelNumber = 'kathtest' OR " & _
        "PersonnelNumber = 'iacs-info' OR " & _
        "PersonnelNumber = 'ee-ad' OR " & _
        "PersonnelNumber = 'plone-Ramanathan-Biophysics' "

        Dim dsSyncData As DataSet = getSQLQueryAsDataset(unwantedQuery)

        Dim tmpRowCount As Integer = dsSyncData.Tables(0).Rows.Count
        Dim queries(tmpRowCount) As String

        For i = 0 To dsSyncData.Tables(0).Rows.Count - 1

            queries(i) = "DELETE FROM EMS_Staging.dbo.tblPeople WHERE PersonnelNumber = '" & dsSyncData.Tables(0).Rows(i).Item(0).ToString.Replace("'", "").Replace("--", "") & "'"

        Next i

        If executeTransactedSQLCommand(queries) = True Then

            Return True

        Else

            Return False

        End If

    End Function


    Public Function MassageDBInfo() As Boolean

        Dim queries(42) As String

        Dim table As String = "EMS_Staging.dbo.tblPeople"

        queries(0) = "UPDATE " & table & " SET City = 'Cambridge', State = 'MA', Zipcode = '02138', Country = 'US' WHERE Phone LIKE '%617%'"
        queries(1) = "UPDATE " & table & " SET Country = 'US' WHERE City = 'Cambridge' AND Country = ''"
        queries(2) = "UPDATE " & table & " SET Address1 = Address2, Address2 = '' WHERE Address1 = '' AND Address2 <> ''"
        queries(3) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, '\n', ''), Address2 = REPLACE(Address2, '\n', '')"
        queries(4) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, '\r', ''), Address2 = REPLACE(Address2, '\r', '')"
        queries(5) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, '\n\r', ''), Address2 = REPLACE(Address2, '\n\r', '')"
        queries(6) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, '\n\r', ''), Address2 = REPLACE(Address2, '\n\r', '')"
        queries(7) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, '""', '')"
        'queries(8) = "DELETE FROM EMS_Staging.dbo." & table & " WHERE NetworkID IN (SELECT NetworkID FROM EMS_Master.dbo.tblWebUser)"
        queries(9) = "DELETE FROM " & table & " WHERE PersonnelNumber = 'dbgtestaccount' OR PersonnelNumber = 'ctaylor' OR PersonnelNumber = 'galileo' OR PersonnelNumber = 'jazkartauser' OR PersonnelNumber = 'jftest45'"
        queries(10) = "DELETE FROM " & table & " WHERE PersonnelNumber = 'rulesuser2' OR PersonnelNumber = 'visitingcommittee' OR PersonnelNumber = 'wwwmatsc' OR PersonnelNumber = 'Win-lap1' OR PersonnelNumber = 'Win-Lap2' OR PersonnelNumber = 'science+cooking'"
        queries(11) = "UPDATE " & table & " SET EMailAddress = PersonnelNumber + '@seas.harvard.edu' WHERE EMailAddress = ''"
        'queries(12) = "UPDATE " & table & " SET EMailAddress = 'ago@seas.harvard.edu' WHERE EmailAddress = 'pirp@seas.harvard.edu' and PersonnelNumber = 'ago'"
        'queries(13) = "UPDATE " & table & " SET EMailAddress = 'jenc@seas.harvard.edu' WHERE PersonnelNumber = 'jenc'"
        queries(14) = "UPDATE " & table & " SET EMailAddress = PersonnelNumber+'@seas.harvard.edu' WHERE EmailAddress NOT LIKE '%seas.harvard.edu'"
        queries(15) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '617/', '(617) ') WHERE Phone LIKE '617/%'"
        queries(16) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '617-', '(617) ') WHERE Phone LIKE '617-%'"
        queries(17) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '617.', '(617) ') WHERE Phone LIKE '617.%'"
        queries(18) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '617 ', '(617) ') WHERE Phone LIKE '617 %'"
        queries(19) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '617', '(617) ') WHERE Phone LIKE '617%' AND Phone NOT LIKE '(617)%'"
        queries(20) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '857/', '(857) ') WHERE Phone LIKE '857/%'"
        queries(21) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '857-', '(857) ') WHERE Phone LIKE '857-%'"
        queries(22) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '857.', '(857) ') WHERE Phone LIKE '857.%'"
        queries(23) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '857 ', '(857) ') WHERE Phone LIKE '857 %'"
        queries(24) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '857', '(857) ') WHERE Phone LIKE '857%' AND Phone NOT LIKE '(857)%'"
        queries(25) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '/', '-') WHERE Phone LIKE '%/%'"
        queries(26) = "UPDATE " & table & " SET Phone = REPLACE(Phone, '.', '-') WHERE Phone LIKE '%/%'"

        queries(27) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '617/', '(617) ') WHERE Fax LIKE '617/%'"
        queries(28) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '617-', '(617) ') WHERE Fax LIKE '617-%'"
        queries(29) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '617.', '(617) ') WHERE Fax LIKE '617.%'"
        queries(30) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '617 ', '(617) ') WHERE Fax LIKE '617 %'"
        queries(31) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '617', '(617) ') WHERE Fax LIKE '617%' AND Fax NOT LIKE '(617)%'"
        queries(32) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '857/', '(857) ') WHERE Fax LIKE '857/%'"
        queries(33) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '857-', '(857) ') WHERE Fax LIKE '857-%'"
        queries(34) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '857.', '(857) ') WHERE Fax LIKE '857.%'"
        queries(35) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '857 ', '(857) ') WHERE Fax LIKE '857 %'"
        queries(36) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '857', '(857) ') WHERE Fax LIKE '857%' AND Fax NOT LIKE '(857)%'"
        queries(37) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '/', '-') WHERE Fax LIKE '%/%'"
        queries(38) = "UPDATE " & table & " SET Fax = REPLACE(Fax, '.', '-') WHERE Fax LIKE '%/%'"

        'queries(39) = "UPDATE " & table & " SET GroupName = GroupID WHERE GroupName = ''"                 'Commented out after DEA modified the HRTK to disable group sync, on our request

        'queries(39) = "UPDATE " & table & " SET BillingReference = NULL"  'Commented out after Nemo HRTK

        queries(40) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, 'Maxwell-Dworkin', 'Maxwell Dworkin'), Address2 = REPLACE(Address2, 'Maxwell-Dworkin', 'Maxwell Dworkin') WHERE Address1 LIKE '%Maxwell-Dworkin%' OR Address2 LIKE '%Maxwell-Dworkin%'"
        queries(41) = "UPDATE " & table & " SET Address1 = REPLACE(Address1, 'Maxwell Dworkin Building', 'Maxwell Dworkin'), Address2 = REPLACE(Address2, 'Maxwell Dworkin Building', 'Maxwell Dworkin') WHERE Address1 LIKE '%Maxwell Dworkin%' OR Address2 LIKE '%Maxwell Dworkin%'"

        If executeTransactedSQLCommand(queries) = True Then

            Return True

        Else

            Return False

        End If

    End Function


    Public Function MassageAddresses() As Boolean

        Dim dsConflictingAddresses As DataSet
        Dim conflictingQueries(62) As String

        conflictingQueries(0) = "SELECT * FROM EMS_Staging.dbo.tblPeople tp WHERE Address1 LIKE '%" & Chr(10) & "%'"
        conflictingQueries(1) = "SELECT * FROM EMS_Staging.dbo.tblPeople tp WHERE Address1 LIKE '%" & Chr(13) & "%'"
        conflictingQueries(2) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp wHERE Address1 LIKE '%29 Oxford St'"
        conflictingQueries(3) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp wHERE Address1 LIKE '%Pierce Hall%' AND Address2 = 'Pierce Hall'"
        conflictingQueries(4) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp wHERE Address1 LIKE '%Pierce%' AND Address2 = 'Pierce Hall'"
        conflictingQueries(5) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%McKay%' AND Address1 NOT LIKE '%Library%'"
        conflictingQueries(6) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%McKay Labs%' AND Address1 NOT LIKE '%Library%'"
        conflictingQueries(7) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Gordon McKay Labs%' AND Address1 NOT LIKE '%Library%'"
        conflictingQueries(8) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%McKay Laboratory%' AND Address1 NOT LIKE '%Library%'"
        conflictingQueries(9) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Cruft%' AND Address1 NOT LIKE '%Cruft Hall%' AND Address1 NOT LIKE '%Cruft Lab%'"
        conflictingQueries(10) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Cruft Lab%' AND Address1 NOT LIKE '%Cruft Hall%'"
        conflictingQueries(11) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Cruft Hall%'"
        conflictingQueries(12) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Maxwell Dworkin%'"
        conflictingQueries(13) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%ESL Lab%'"
        conflictingQueries(14) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Engineering Sciences Laboratory%'"
        conflictingQueries(15) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%60 Oxford St.%'"
        conflictingQueries(16) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%60 Oxford Street%'"
        conflictingQueries(17) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%LISE%'"
        conflictingQueries(18) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%NWB%'"
        conflictingQueries(19) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%NW%'"
        conflictingQueries(20) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%North West Building%'"
        conflictingQueries(21) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address2 LIKE '%North West Building%'"
        conflictingQueries(22) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE 'Pierce'"
        conflictingQueries(23) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE 'Pierce Hall'"
        conflictingQueries(24) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%School of Engineering and Applied Sciences%'"
        conflictingQueries(25) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%SEAS%'"
        conflictingQueries(26) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Engineering Sciences Lab%'"
        conflictingQueries(27) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%McKay%' AND Address2 = 'McKay Laboratory of Applied Science'"
        conflictingQueries(28) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Geological Museum%'"
        conflictingQueries(29) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address2 LIKE '%Geological Museum%' AND Address1 LIKE '%Museum%'"
        conflictingQueries(30) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Jefferson Lab%'"
        conflictingQueries(31) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Jefferson%'"
        conflictingQueries(32) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Hoffman%'"
        conflictingQueries(33) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%iLab%'"
        conflictingQueries(34) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Mallinckrodt Chemistry Lab%'"
        conflictingQueries(35) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Mallinckrodt%'"
        conflictingQueries(36) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Lyman%'"
        conflictingQueries(37) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address2 = 'Mallinckrodt'"
        conflictingQueries(38) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Harvard University, %'"
        conflictingQueries(39) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Harvard%'"
        conflictingQueries(40) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 = '9 Oxford St' AND Address2 = ''"
        conflictingQueries(41) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 = '29 Oxford St' AND Address2 = ''"
        conflictingQueries(42) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 = '33 Oxford St' AND Address2 = ''"
        conflictingQueries(43) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%29 Oxford St%Pierce Hall%'"
        conflictingQueries(44) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%29 Oxford St%15 Oxford St%'"
        conflictingQueries(45) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%29 Oxford St%9 Oxford Street%'"
        conflictingQueries(45) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%29 Oxford St%9 Oxford Street%'"
        conflictingQueries(46) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Cambridge, MA, 02138%'"
        conflictingQueries(47) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%60 Oxford St%' AND Address2 LIKE '%60 Oxford St%'"
        conflictingQueries(48) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%29 Oxford Street%'"
        conflictingQueries(49) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%29 Oxford St%29 Oxford St%'"
        conflictingQueries(50) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%52 Oxford St%52 Oxford St%'"
        conflictingQueries(51) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Rm. %'"
        conflictingQueries(52) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Rm %'"
        conflictingQueries(53) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Room %'"
        conflictingQueries(54) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%,%'"
        conflictingQueries(55) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Link%'"
        conflictingQueries(56) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Oxford%'"
        conflictingQueries(57) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Cambridge MA 02138%'"
        conflictingQueries(58) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%Street%'"
        conflictingQueries(59) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%#%'"
        conflictingQueries(60) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%12 Oxford St%12 Oxford St%'"
        conflictingQueries(61) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Address1 LIKE '%11 Oxford St%11 Oxford St%'"
        conflictingQueries(62) = "SELECT * FROM [EMS_Staging].[dbo].[tblPeople] tp WHERE Country = 'US'"

        For i = 0 To conflictingQueries.Length - 1

            dsConflictingAddresses = getSQLQueryAsDataset(conflictingQueries(i))

            Dim numberOfConflictingRows As Integer = dsConflictingAddresses.Tables(0).Rows.Count

            If numberOfConflictingRows > 0 Then

                If i = 0 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace(Chr(10), "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 1 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace(Chr(13), "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 2 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Pierce Hall', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("29 Oxford St", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 3 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Pierce Hall", "29 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 4 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Pierce", "29 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 5 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='McKay Laboratory of Applied Science' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 6 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='McKay Laboratory of Applied Science', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("McKay Labs", "9 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 7 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='McKay Laboratory of Applied Science', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Gordon McKay Labs", "9 Oxford St").Replace("15 Oxford St", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 8 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='McKay Laboratory of Applied Science', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("McKay Laboratory", "9 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 9 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Cruft Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("19 Oxford St", "").Replace("Cruft", "19A Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 10 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Cruft Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("19 Oxford St", "").Replace("Cruft Lab", "19A Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 11 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Cruft Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("19 Oxford St", "").Replace("Cruft Hall", "19A Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 12 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Maxwell Dworkin', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("33 Oxford St", "").Replace("Maxwell Dworkin", "33 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 13 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Engineering Sciences Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("40 Oxford St", "").Replace("ESL Lab", "58 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 14 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Engineering Sciences Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Engineering Sciences Laboratory", "58 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 15 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("60 Oxford St.", "60 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 16 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("60 Oxford Street", "60 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 17 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='LISE', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("LISE", "15 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 18 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Northwest Building', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("NWB", "52 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 19 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Northwest Building', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("NW", "52 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 20 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Northwest Building', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("North West Building", "52 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 21 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Northwest Building', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("North West Building", "52 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 22 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Pierce Hall', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Pierce", "29 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 23 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Pierce Hall', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Pierce Hall", "29 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 24 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("School of Engineering and Applied Sciences", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 25 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("SEAS", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 26 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Engineering Sciences Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Engineering Sciences Lab", "58 Oxford St").Replace("40 Oxford St", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 27 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("McKay", "9 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 28 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Geological Museum', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Geological Museum", "26 Oxford St").Replace("20 Oxford St", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 29 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Museum", "").Replace("20 Oxford St", "26 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 30 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Jefferson Physical Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("17 Oxford St", "").Replace("Jefferson Lab", "17 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 31 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Jefferson Physical Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("17 Oxford St", "").Replace("Jefferson", "17 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 32 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Hoffman Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Hoffman", "20 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 33 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='iLab', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("iLab", "125 Western Ave").Trim & "', City = 'Allston', State = 'MA', Country = 'USA', Zipcode = '02163' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 34 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Mallinckrodt Chemistry Lab', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Mallinckrodt Chemistry Lab", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 35 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Mallinckrodt Chemistry Lab', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Mallinckrodt", "12 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 36 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Lyman Laboratory', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Lyman", "11 Oxford St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 37 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Mallinckrodt Chemistry Lab' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 38 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Harvard University, ", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 39 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Harvard", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 40 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='McKay Laboratory of Applied Science' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 41 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Pierce Hall' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 42 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 ='Maxwell Dworkin' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 43 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Pierce Hall', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Pierce Hall", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 44 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Pierce Hall', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("15 Oxford St", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 45 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Pierce Hall', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("9 Oxford St", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 46 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Cambridge, MA, 02138", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 47 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = '' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 48 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("29 Oxford Street", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 49 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String
                    Dim tmpValue As String

                    For j = 0 To numberOfConflictingRows - 1

                        tmpValue = dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(0, dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("29 Oxford St")) & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("29 Oxford St") + 12)
                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Pierce Hall', Address1 = '" & tmpValue.Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 50 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String
                    Dim tmpValue As String

                    For j = 0 To numberOfConflictingRows - 1

                        tmpValue = dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(0, dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("52 Oxford St")) & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("52 Oxford St") + 12)
                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Northwest Building', Address1 = '" & tmpValue.Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 51 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Rm. ", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 52 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Rm ", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 53 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Room ", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 54 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace(",", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 55 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Mallinckrodt Chemistry Lab', Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Link Building", "").Replace("Mallinckrodt Link", "").Replace("Link", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 56 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET City = 'Cambridge', State = 'MA', ZipCode = '02138', Country = 'USA' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 57 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Cambridge MA 02138", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 58 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("Street", "St").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 59 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Trim.Replace("#", "").Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 60 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String
                    Dim tmpValue As String

                    For j = 0 To numberOfConflictingRows - 1

                        tmpValue = dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(0, dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("12 Oxford St")) & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("12 Oxford St") + 12)
                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address2 = 'Mallinckrodt Chemistry Lab', Address1 = '" & tmpValue.Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 61 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String
                    Dim tmpValue As String

                    For j = 0 To numberOfConflictingRows - 1

                        tmpValue = dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(0, dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("11 Oxford St")) & dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.Substring(dsConflictingAddresses.Tables(0).Rows(j).Item("Address1").ToString.LastIndexOf("11 Oxford St") + 12)
                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Address1 = '" & tmpValue.Trim & "' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)

                ElseIf i = 62 Then

                    Dim resolvingQueries(numberOfConflictingRows) As String

                    For j = 0 To numberOfConflictingRows - 1

                        resolvingQueries(j) = "UPDATE EMS_Staging.dbo.tblPeople SET Country = 'USA' WHERE PersonnelNumber = '" & dsConflictingAddresses.Tables(0).Rows(j).Item("PersonnelNumber").ToString & "'"

                    Next j

                    executeTransactedSQLCommand(resolvingQueries)


                End If

            End If

        Next i

        Return True

    End Function


    'Public Function ConsolidateEMSGroups() As Boolean

    '    Dim dsGroups As DataSet
    '    Dim dsMembersCountTypes As DataSet

    '    dsGroups = getSQLQueryAsDataset("SELECT DISTINCT GroupID FROM EMS_Staging.dbo.tblPeople ORDER BY 1 ASC")

    '    For i = 0 To dsGroups.Tables(0).Rows.Count - 1

    '        dsMembersCountTypes = getSQLQueryAsDataset("SELECT GroupID, GroupType, COUNT(*) AS groupMembersCount FROM EMS_Staging.dbo.tblPeople WHERE GroupID = '" & dsGroups.Tables(0).Rows(i).Item(0) & "' GROUP BY GroupID, GroupType ORDER BY 1 ASC, 3 DESC")

    '        If dsMembersCountTypes.Tables(0).Rows.Count > 1 Then 'When it has at least two types

    '            executeSQLCommand("UPDATE EMS_Staging.dbo.tblPeople SET GroupType = '" & dsMembersCountTypes.Tables(0).Rows(0).Item(1) & "' WHERE GroupID = '" & dsGroups.Tables(0).Rows(i).Item(0) & "'")

    '        End If

    '    Next i

    '    Return True

    'End Function


    Public Function PutUsersInDefaultGroups() As Boolean

        Return executeSQLCommand("INSERT INTO [EMS_Staging].[dbo].[tblGroupMemberships] SELECT tp.PersonnelNumber, 34 FROM [EMS_Staging].[dbo].[tblPeople] tp LEFT JOIN [EMS_Staging].[dbo].[tblGroupMemberships] tgm ON tp.PersonnelNumber = tgm.PersonnelNumber WHERE tgm.GroupId IS NULL ORDER BY 2 ASC, 1 ASC")

    End Function


    Public Function runEMSStoreProcedure() As Boolean

        If executeSQLCommand("exec EMS.dbo.HRTK_Update_Group") = True Then

            Return True

        Else

            Return False

        End If

    End Function


    Public Function SyncADUserInfoToEMSUsers() As Boolean

        SyncUsers()
        DeleteInactiveUsers()
        DeleteUnwantedUsers()
        MassageDBInfo()
        MassageAddresses()
        PutUsersInDefaultGroups()   'ConsolidateEMSGroups()
        'SyncUserDelegates()
        runEMSStoreProcedure()

        Return True

    End Function


End Class
