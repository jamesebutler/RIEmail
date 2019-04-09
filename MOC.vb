
Imports System.Reflection

Module MOC
    Sub GetMOCEnteredEmails(ByVal dtRunDate)
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim dr, drEmail As DataRow
        Dim dsSites As DataSet

        Dim strErr As String
        Dim strSiteId As String = ""
        Dim strEmailAddress As String = ""
        Dim param As New OracleParameter
        Dim intSiteCnt, intEmailCnt, intEmailSent As Integer

        Try
            'Set up db connection
            'strSP = My.Application.Info.AssemblyName
            InsertAuditRecord(strAppName, strSP & " started")

            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            'Set up initial procedure which returns the profile for emails that should be sent based on run date passed in.
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MOCBATCHEMAILS.SITELISTING"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_date"
                param.OracleDbType = OracleDbType.Date
                param.Direction = Data.ParameterDirection.Input
                param.Value = dtRunDate
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsSites"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            Dim daSites = New OracleDataAdapter(cmdSQL)
            dsSites = New Data.DataSet()
            daSites.Fill(dsSites)

            'Get Site count so we can determine if all emails were sent
            intSiteCnt = dsSites.Tables(0).Rows.Count()

            'Loop thru records for the DEFAULT Profile
            For Each dr In dsSites.Tables(0).Rows

                strSiteId = dr("Siteid")

                Try
                    cmdSQL = New OracleCommand
                    With cmdSQL
                        .Connection = connDB
                        .CommandText = "MOCBATCHEMAILS.EMailLISTING"
                        .CommandType = CommandType.StoredProcedure

                        param = New OracleParameter
                        param.ParameterName = "in_date"
                        param.OracleDbType = OracleDbType.Date
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = dtRunDate
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "in_siteid"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strSiteId
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "rsUserIds"
                        param.OracleDbType = OracleDbType.Cursor
                        param.Direction = ParameterDirection.Output
                        .Parameters.Add(param)

                    End With

                    Dim daEmail = New OracleDataAdapter(cmdSQL)
                    Dim dsEmail As DataSet
                    dsEmail = New DataSet()
                    daEmail.Fill(dsEmail)

                    Dim i As Integer = 0
                    'Loop thru all email records 
                    For Each drEmail In dsEmail.Tables(0).Rows
                        i = i + 1
                        If i = 1 Then
                            strEmailAddress = drEmail("email")
                        Else
                            strEmailAddress = strEmailAddress & "," & drEmail("email")
                        End If

                    Next
                    GetMOCs(strSiteId, dtRunDate, strEmailAddress)

                    strEmailAddress = ""
                    drEmail = Nothing
                    dsEmail = Nothing

                Catch ex As Exception
                    'Exception handling
                    strErr = "Error occurred. " & ex.Message
                    SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", strErr)
                    HandleError(strSP, strErr, ex)
                End Try

                intEmailSent = intEmailSent + 1

            Next

            'If counts indicate not all emails were sent, send email to support and write record to audit table.
            If intEmailSent <> intSiteCnt Then
                SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", "Only " & intEmailSent & " emails were sent.  " & intEmailCnt & " Emails should have been sent.")
                HandleError(strSP, "Only " & intEmailSent & " emails were sent.  " & intSiteCnt & " Emails should have been sent")
            Else
                InsertAuditRecord(strAppName & "." & strSP, intEmailSent & " of " & intSiteCnt & " emails were sent.")
            End If

            dsSites = Nothing
            daSites = Nothing
            connDB.Close()

            InsertAuditRecord(strAppName, strSP & " ended")

        Catch ex As Exception
            HandleError(strSP, ex.Message, ex)
        Finally
            connDB.Close()
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub

    Sub GetMOCs(ByVal strSiteid As String, ByVal dtRunDate As Date, ByVal strEmailAddress As String)
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drMOCs As OracleDataReader
        Dim strMOCID, strMOCCategory, strMOCClassification, strInitiator As String
        Dim strArea, strMOCType, strStatus, strTitle, strImplementationDate, strExpirationDate As String

        Dim strHeading As String = ""
        Dim strHeading1 As String = ""
        Dim strErr As String
        Dim strSubject As String = ""
        Dim strMsg As String
        Dim strBody As String = ""
        Dim strFooter As String = ""
        Dim previous_recType, strRecType
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder
        Dim strTasksFound As String = "N"

        Try
            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            previous_recType = ""
            strSP = My.Application.Info.AssemblyName

            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MOCBATCHEMAILS.MOCLISTING"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_date"
                param.OracleDbType = OracleDbType.Date
                param.Direction = Data.ParameterDirection.Input
                param.Value = dtRunDate
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "in_siteid"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strSiteid
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsAllMOCs"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            sbEmailBody = New System.Text.StringBuilder
            Dim v_td As String() = {"<TD>", "</TD>"}

            drMOCs = cmdSQL.ExecuteReader()

            While drMOCs.Read
                strRecType = drMOCs("RecType")

                strArea = drMOCs("RIArea")
                strMOCType = drMOCs("MOCType")
                If drMOCs("MOCCategory") Is DBNull.Value Then
                    strMOCCategory = ""
                Else
                    strMOCCategory = drMOCs("MOCCategory")
                End If
                strTitle = drMOCs("Title")
                If drMOCs("MOCClassification") Is DBNull.Value Then
                    strMOCClassification = ""
                Else
                    strMOCClassification = drMOCs("MOCClassification")
                End If
                strMOCID = drMOCs("MOCNumber")
                strStatus = drMOCs("Status")
                strImplementationDate = drMOCs("StartDate") 'Cannot be NULL
                strExpirationDate = drMOCs("EndDate") 'Cannot be NULL
                strInitiator = drMOCs("person")

                If previous_recType <> strRecType Or previous_recType = "" Then
                    sbEmailBody.Append("</table><P><font size=2 face=Arial><B><U>" & strRecType & "</B></U></FONT><BR>")
                    sbEmailBody.Append("<TABLE border=1><TR valign=top><font size=2 face=Arial><B><TD width=30%> Title{1}<TD width=10%>Status{1}<TD width=10%>Implementation Date{1}")
                    sbEmailBody.Append("{0}Initiator{1}<TD width=15%>Area{1}{0}Type{1}{0}Category{1}{0}Classification{1}")
                    sbEmailBody.Append("</B></TR>")
                End If

                sbEmailBody.Append("<BR><TR valign=top><font size=2 face=Arial>")
                sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}")
                sbEmailBody.Append("{0}" & strStatus & "{1}")
                sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                sbEmailBody.Append("{0}" & strInitiator & "{1}")
                sbEmailBody.Append("{0}" & strArea & "{1}")
                sbEmailBody.Append("{0}" & strMOCType & "{1}")
                sbEmailBody.Append("{0}" & strMOCCategory & "{1}")
                sbEmailBody.Append("{0}" & strMOCClassification & "{1}")

                previous_recType = strRecType
            End While

            strMsg = sbEmailBody.ToString
            strMsg = String.Format(strMsg, v_td)
            strSubject = "MOC - Entered/Updated in the last week"
            If strDB = "RIDEV" Then
                strHeading1 = "***THIS IS A TEST NOTIFICATION***"
            End If
            strHeading = "<HTML><BODY><font size=2 face=Arial><B>Following are all MOC records that have been added or updated in the past week.<BR><BR>Click on MOC Title to transfer to MOC update screen.</B>"
            strFooter = "</HTML></BODY>"

            If dtRunDate <> Now().Date Then
                strBody = strHeading1 & "<P><font size =1 face=Arial><B>MOC BATCH EMAIL RERUN for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR><font size=1>" & strHeading & "<BR>" & strMsg.ToString & strFooter
            Else
                'strBody = "<P><font size =1 face=Arial><B>MTT BATCH EMAIL for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR>" & strHeading & "<BR>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                strBody = strHeading1 & "<P><font size =1 face=Arial><B>" & strHeading & "<BR>" & strMsg.ToString & strFooter
            End If

            strBody = cleanString(strBody, "<br>")

            SendEmail(strEmailAddress, "RootCause.FailureAnalysis@graphicpkg.com", strSubject, strBody, "MOC")

            strBody = ""

            drMOCs = Nothing

        Catch ex As Exception
            'Exception handling
            strErr = "Error occurred." & ex.Message
            SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", strErr)
            HandleError(strSP, strErr, ex)
        Finally
            connDB.Close()
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub


    Sub GetTempMOCUserIds()
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim dr As DataRow
        Dim dsUserID As DataSet

        Dim strErr As String
        Dim strUserId As String
        Dim param As New OracleParameter
        Dim intUserIDCnt, intEmailSent As Integer
        Dim sbEmailBody As New System.Text.StringBuilder
        Dim strLanguage As String = "en-US"

        Try
            InsertAuditRecord(strAppName, strSP & " started")

            'Set up db connection
            strSP = My.Application.Info.AssemblyName

            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            'Set up initial procedure which returns the profile for emails that should be sent based on run date passed in.
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MOCBATCHEMAILS.TEMPMOCEMAILS"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "rsTEMPMOCEMails"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            Dim daUserID = New OracleDataAdapter(cmdSQL)
            dsUserID = New DataSet()
            daUserID.Fill(dsUserID)

            'Get Site count so we can determine if all emails were sent
            intUserIDCnt = dsUserID.Tables(0).Rows.Count()

            'Loop thru records for the DEFAULT Profile
            For Each dr In dsUserID.Tables(0).Rows

                strUserId = dr("UserName")

                If dr("default_language") <> "" Then
                    strLanguage = dr("default_language")
                Else
                    strLanguage = "en-US"
                End If

                Try
                    GetTempMOC(strUserId, strLanguage)

                Catch ex As Exception
                    'Exception handling
                    strErr = "Error occurred." & ex.Message
                    SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", strErr)
                    HandleError(strSP, strErr, ex)
                End Try

                intEmailSent = intEmailSent + 1

            Next

            'If counts indicate not all emails were sent, send email to support and write record to audit table.
            If intEmailSent <> intUserIDCnt Then
                SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", "Only " & intEmailSent & " emails were sent.  " & intUserIDCnt & " Emails should have been sent.")
                HandleError(strSP, "Only " & intEmailSent & " emails were sent.  " & intUserIDCnt & " Emails should have been sent")
            Else
                InsertAuditRecord(strAppName & "." & strSP, intEmailSent & " of " & intUserIDCnt & " emails were sent.")
            End If

            dsUserID = Nothing
            daUserID = Nothing

            InsertAuditRecord(strAppName, strSP & " ended")

        Catch ex As Exception
            HandleError(strSP, ex.Message, ex)
        Finally
            connDB.Close()
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub

    Sub GetTempMOC(ByVal strUserid As String, ByVal strLanguage As String)
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drMOCs As OracleDataReader
        Dim strMOCID As String
        Dim strEmailAddress As String = ""
        Dim strStatus, strTitle, strImplementationDate, strExpirationDate, strErr As String

        Dim strHeading As String = ""
        Dim strHeading1 As String = ""
        Dim strSubject As String = ""
        Dim strMsg As String = ""
        Dim strBody As String = ""
        Dim strFooter As String = ""
        Dim previous_recType As String = ""
        Dim strRecType As String
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder

        Try
            Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")



            'Try
            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MOCBATCHEMAILS.TEMPMOCLISTING"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_userid"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strUserid
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsTEMPMOCs"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            sbEmailBody = New System.Text.StringBuilder
            Dim v_td As String() = {"<TD>", "</TD>"}

            drMOCs = cmdSQL.ExecuteReader()

            While drMOCs.Read
                strRecType = drMOCs("RecType")
                strMOCID = drMOCs("MOCNumber")
                strTitle = drMOCs("Incident")
                strStatus = IPLoc.GetResourceValue(drMOCs("Status"))
                strImplementationDate = drMOCs("StartDate") 'Cannot be NULL
                strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
                strExpirationDate = drMOCs("EndDate") 'Cannot be NULL
                strExpirationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strExpirationDate, strLanguage, "dd MMM yyyy")
                strEmailAddress = drMOCs("Email")

                If previous_recType <> strRecType Or previous_recType = "" Then
                    sbEmailBody.Append("</table><BR>")
                    sbEmailBody.Append("<TABLE border=1><TR valign=top><font size=2 face=Arial><B><TD width=40%>" & IPLoc.GetResourceValue("Title") & "{1}<TD width=15%>" & IPLoc.GetResourceValue("Status") & "{1}<TD width=15%>" & IPLoc.GetResourceValue("Implementation Date") & "{1}")
                    sbEmailBody.Append("{0}" & IPLoc.GetResourceValue("Expiration Date") & "{1}")
                    sbEmailBody.Append("</B></TR>")
                End If

                sbEmailBody.Append("<BR><TR valign=top><font size=2 face=Arial>")
                sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}")
                sbEmailBody.Append("{0}" & strStatus & "{1}")
                sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                sbEmailBody.Append("{0}" & strExpirationDate & "{1}")

                previous_recType = strRecType
            End While

            strMsg = sbEmailBody.ToString
            strMsg = String.Format(strMsg, v_td)
            strSubject = IPLoc.GetResourceValue("MOC - Temporary Records")
            If strDB = "RIDEV" Then
                strHeading1 = "***THIS IS A TEST NOTIFICATION***"
            End If
            strHeading = "<HTML><BODY><font size=2 face=Arial><B>" & IPLoc.GetResourceValue("The following are Trial/Temporary Management of Change records that were initiated by you.<BR><BR>Click on MOC Title to transfer to MOC update screen.") & "</B>"
            strHeading = strHeading & "<B>  " & IPLoc.GetResourceValue("Please review to make sure they should still be designated as Trial/Temporary.</B>")
            strFooter = "</HTML></BODY>"

            If dtRunDate <> Now().Date Then
                strBody = strHeading1 & "<P><font size =1 face=Arial><B>MOC BATCH EMAIL RERUN for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR><font size=1>" & strHeading & "<BR>" & strMsg.ToString & strFooter
            Else
                'strBody = "<P><font size =1 face=Arial><B>MTT BATCH EMAIL for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR>" & strHeading & "<BR>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                strBody = strHeading1 & "<P><font size =1 face=Arial><B>" & strHeading & "<BR>" & strMsg.ToString & strFooter
            End If

            strBody = cleanString(strBody, "<br>")

            SendEmail(strEmailAddress, "RootCause.FailureAnalysis@graphicpkg.com", strSubject, strBody, "MOC")

            strBody = ""

            drMOCs = Nothing

        Catch ex As Exception
            'Exception handling
            strErr = "Error occurred." & ex.Message
            SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", strErr, "MOC")
            HandleError(strSP, strErr, ex)
        Finally
            connDB.Close()
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub
    Sub GetPendingMOCUserids()
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim dr As DataRow
        Dim dsUserID As DataSet

        Dim strErr As String
        Dim strUserId As String
        Dim strLanguage As String
        Dim strEmail As String
        Dim param As New OracleParameter
        Dim intUserIDCnt, intEmailSent As Integer

        Try
            'Set up db connection
            'strSP = My.Application.Info.AssemblyName

            InsertAuditRecord(strAppName, strSP & " started")

            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            'Set up initial procedure which returns the profile for emails that should be sent based on run date passed in.
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MOCBATCHEMAILS.PENDINGApprovalEMAILS"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "rsApprovalEMails"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            Dim daUserID = New OracleDataAdapter(cmdSQL)
            dsUserID = New DataSet()
            daUserID.Fill(dsUserID)

            'Get Site count so we can determine if all emails were sent
            intUserIDCnt = dsUserID.Tables(0).Rows.Count()

            'Loop thru records for the DEFAULT Profile
            For Each dr In dsUserID.Tables(0).Rows

                strUserId = dr("UserName")
                strEmail = dr("Email")

                If dr("default_language") <> "" Then
                    strLanguage = dr("default_language")
                Else
                    strLanguage = "en-US"
                End If

                Try
                    GetPendingMOCs(strUserId, strLanguage, strEmail)

                Catch ex As Exception
                    'Exception handling
                    strErr = "Error occurred." & ex.Message
                    SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", strErr)
                    HandleError(strSP, strErr, ex)
                End Try

                intEmailSent = intEmailSent + 1

            Next

            dsUserID = Nothing
            daUserID = Nothing

            InsertAuditRecord(strAppName, strSP & " ended")

        Catch ex As Exception
            HandleError(strSP, ex.Message, ex)
        Finally
            connDB.Close()
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub

    'Sub GetPendingMOCs(ByVal strUserid As String, ByVal strLanguage As String)
    '    'ALA    1/19/2017 
    '    '       Add Initiator to pending MOC email

    '    Dim connDB As New OracleConnection
    '    Dim cmdSQL As OracleCommand = Nothing
    '    Dim drMOCs As OracleDataReader
    '    Dim strMOCID As String
    '    Dim strEmailAddress As String = ""
    '    Dim strTitle, strImplementationDate, strErr As String
    '    Dim strInitiator As String

    '    Dim strHeading As String = ""
    '    Dim strHeading1 As String = ""
    '    Dim strSubject As String = ""
    '    Dim strMsg As String = ""
    '    Dim strDesc As String = ""
    '    Dim strBody As String = ""
    '    Dim strFooter As String = ""
    '    Dim previous_recType As String = ""
    '    Dim strRecType As String
    '    Dim param As New OracleParameter
    '    Dim sbEmailBody As New System.Text.StringBuilder


    '    Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")
    '    'Dim ip As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")

    '    'IPLoc.ClearResourceCache()

    '    Try
    '        If strDB = "RIDEV" Then
    '            connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
    '        Else
    '            connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
    '        End If
    '        connDB.Open()

    '        cmdSQL = New OracleCommand
    '        With cmdSQL
    '            .Connection = connDB
    '            .CommandText = "MOCBATCHEMAILS.PENDINGMOCLISTING"
    '            .CommandType = CommandType.StoredProcedure

    '            param = New OracleParameter
    '            param.ParameterName = "in_username"
    '            param.OracleDbType = OracleDbType.VarChar
    '            param.Direction = Data.ParameterDirection.Input
    '            param.Value = strUserid
    '            .Parameters.Add(param)

    '            param = New OracleParameter
    '            param.ParameterName = "rsPendingMOCs"
    '            param.OracleDbType = OracleDbType.Cursor
    '            param.Direction = ParameterDirection.Output
    '            .Parameters.Add(param)

    '        End With

    '        sbEmailBody = New System.Text.StringBuilder
    '        Dim v_td As String() = {"<TD>", "</TD>"}

    '        drMOCs = cmdSQL.ExecuteReader()
    '        Dim i As Integer = 1
    '        While drMOCs.Read
    '            strRecType = drMOCs("RecType")
    '            strMOCID = drMOCs("MOCNumber")
    '            strTitle = drMOCs("Incident")
    '            'strSite = drMOCs("Siteid")
    '            If drMOCs("Description") Is DBNull.Value Then
    '                strDesc = ""
    '            Else
    '                strDesc = drMOCs("Description")
    '                strDesc = Replace(strDesc, "{", "(")
    '                strDesc = Replace(strDesc, "}", ")")
    '            End If
    '            'strStatus = drMOCs("Status")
    '            strImplementationDate = drMOCs("EventDate") 'Cannot be NULL
    '            strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
    '            'strExpirationDate = drMOCs("EndDate") 'Cannot be NULL
    '            strInitiator = drMOCs("Initiator")
    '            strEmailAddress = drMOCs("Email")

    '            If previous_recType <> strRecType Or previous_recType = "" Then
    '                'If i = 1 Then
    '                sbEmailBody.Append("</table><P><font size=2 face=Arial><B><U>" & IPLoc.GetResourceValue(strRecType) & "</B></U></FONT><BR>")
    '                'sbEmailBody.Append("<BR>" & strRecType & "</table><BR>")
    '                sbEmailBody.Append("<TABLE border=1><TR valign=top><font size=2 face=Arial><B><TD width=20%> " & IPLoc.GetResourceValue("Title") & "{1}<TD width=40%> " & IPLoc.GetResourceValue("Description") & "{1}<TD width=15%> " & IPLoc.GetResourceValue("Initiator") & "{1}<TD width=10%>" & IPLoc.GetResourceValue("Implementation Date") & "{1}")
    '                'sbEmailBody.Append("<TABLE border=1><TR valign=top><font size=2 face=Arial><B><TD width=20%> " & IPLoc.GetResourceValue("Title") & "{1}<TD width=45%> " & IPLoc.GetResourceValue("Description") & "{1}<TD width=15%>" & IPLoc.GetResourceValue("Implementation Date") & "{1}")
    '                sbEmailBody.Append("</B></TR>")
    '            End If

    '            sbEmailBody.Append("<BR><TR valign=top><font size=2 face=Arial>")
    '            sbEmailBody.Append("{0}<A HREF=HTTP://" & strDB & strURL & "/RI/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}")
    '            sbEmailBody.Append("{0}" & strDesc & "{1}")
    '            sbEmailBody.Append("{0}" & strInitiator & "{1}")
    '            sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
    '            If strRecType = "Pending Your Review" Then
    '                sbEmailBody.Append("{0}<A HREF=HTTP://" & strDB & strURL & "/RI/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & "&UserName=" & strUserid & ">" & IPLoc.GetResourceValue("Click the blue link to MOC to acknowledge that you have reviewed the MOC") & "</A>{1}")
    '            End If
    '            'sbEmailBody.Append("{0}" & strExpirationDate & "{1}")

    '            previous_recType = strRecType
    '            i = i + 1
    '        End While

    '        strMsg = sbEmailBody.ToString
    '        strMsg = String.Format(strMsg, v_td)
    '        strSubject = IPLoc.GetResourceValue("MOCs - Pending Approval or Review")
    '        If strDB = "RIDEV" Then
    '            strHeading1 = "***THIS IS A TEST NOTIFICATION*** MOC BATCH EMAIL RERUN for " & dtRunDate & " " & strEmailAddress
    '            strEmailAddress = "amy.albrinck@ipaper.com"
    '        End If
    '        strHeading = "<HTML><BODY><font size=2 face=Arial><B>" & IPLoc.GetResourceValue("The following are MOCs that are pending your approval or review.<BR><BR>Click on Title to update MOC.") & "</B>"
    '        strHeading = strHeading & "<B> " & IPLoc.GetResourceValue("Please review.") & "</B><br>"
    '        strFooter = "</HTML></BODY>"

    '        If dtRunDate <> Now().Date Then
    '            strBody = strHeading1 & "<P><font size =1 face=Arial><B>MOC BATCH EMAIL RERUN for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR><font size=1>" & strHeading & "<BR>" & strMsg.ToString & strFooter
    '        Else
    '            'strBody = "<P><font size =1 face=Arial><B>MTT BATCH EMAIL for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR>" & strHeading & "<BR>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
    '            strBody = strHeading1 & "<P><font size =1 face=Arial><B>" & strHeading & "<BR>" & strMsg.ToString & strFooter
    '        End If

    '        strBody = cleanString(strBody, "<br>")

    '        SendEmail(strEmailAddress, "RootCause.FailureAnalysis@ipaper.com", strSubject, strBody, "MOC")

    '        strBody = ""

    '        drMOCs = Nothing

    '    Catch ex As Exception
    '        'Exception handling
    '        strErr = "Error occurred." & ex.Message
    '        SendEmail(supportEmail, "RootCause.FailureAnalysis@ipaper.com", "RIBatchEmail", strErr)
    '        HandleError(strSP, strErr, ex)
    '    Finally
    '        connDB.Close()
    '        If Not connDB Is Nothing Then connDB = Nothing
    '        If Not cmdSQL Is Nothing Then cmdSQL = Nothing
    '    End Try
    'End Sub

    Sub GetPendingMOCs(ByVal strUserid As String, ByVal strLanguage As String, ByVal strEmailAddress As String)
        'ALA    1/19/2017 
        '       Add Initiator to pending MOC email
        '
        'ALA    4/17/2017
        '       Change to use MOCMyMOCS.GetMYMOCs

        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drMOCs As OracleDataReader

        Dim strMOCID As String
        'Dim strEmailAddress As String = ""
        Dim strTitle, strImplementationDate, strErr As String
        Dim strInitiator As String

        Dim strHeading As String = ""
        Dim strHeading1 As String = ""
        Dim strSubject As String = ""
        Dim strMsg As String = ""
        Dim strDesc As String = ""
        Dim strBody As String = ""
        Dim strFooter As String = ""
        Dim previous_recType As String = ""
        Dim strRecType As String
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder

        'strUserid = "AALBRIN"
        Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")

        Try
            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MOCMYMOCS.GetMYMOCs"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_username"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strUserid
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsMOCs"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsMOCDrafts"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsMOCOnHold"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsMOCImplOverride"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsMOCCompOverride"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsOwnerPendingMOCs"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsApprovedNotImplementedMOCs"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            sbEmailBody = New System.Text.StringBuilder
            Dim v_td As String() = {"<TD style='border: solid 1px #DDEEEE;  color: #333;  padding: 0px; border-width:0px; margin:0px; text-shadow: 1px 1px 1px #fff;'>", "</TD>"}
            Dim v_table As String = "<TABLE width='90%' style='border: 1px solid #DDD;border-collapse: collapse; border-spacing:0px; font: normal 13px Arial, sans-serif;'>"
            Dim v_th As String = "<TH style='background-color: #DDEFEF;    border: solid 1px #DDEEEE;    color: #336B6B;    text-align: left; vertical-align:top;'>"
            'Dim v_td As String() = "<TD style={border: solid 1px #DDEEEE;color: #333;padding: 10px;text-shadow: 1px 1px 1px #fff;>", "</TD>"}

            'strHeading = "<HTML><BODY><B><font size=2 face=Arial>" & IPLoc.GetResourceValue("The following are MOCs that are pending your approval or review.<BR><BR>Click on Title to update MOC.") & "</B>"
            'strHeading = strHeading & "<B> " & IPLoc.GetResourceValue("Please review.") & "</B><br>"

            drMOCs = cmdSQL.ExecuteReader()
            Dim i As Integer = 1
            While drMOCs.Read
                strRecType = drMOCs("approval_Type")
                strMOCID = drMOCs("MOCNumber")
                strTitle = drMOCs("title")
                'strSite = drMOCs("Siteid")
                If drMOCs("Description") Is DBNull.Value Then
                    strDesc = ""
                Else
                    strDesc = drMOCs("Description")
                    strDesc = Replace(strDesc, "{", "(")
                    strDesc = Replace(strDesc, "}", ")")
                End If

                'strStatus = drMOCs("Status")
                If drMOCs("EventDate") Is DBNull.Value Then
                    strImplementationDate = String.Empty
                Else
                    strImplementationDate = drMOCs("EventDate")
                    strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
                End If
                strInitiator = drMOCs("initiatorname")
                'strEmailAddress = drMOCs("Email")

                If previous_recType <> strRecType Or previous_recType = "" Then
                    sbEmailBody.Append("</table><P style='    padding: 10px; background-color: #e0e0eb;    border: solid 1px #DDEEEE;    color: #336B6B;'><font size=2 face=Arial><B><U>" & IPLoc.GetResourceValue(strRecType) & "</B></U></FONT>")
                    sbEmailBody.Append(v_table & "<thead><TR><B>" & v_th & IPLoc.GetResourceValue("Title") & "</TH>" & v_th & IPLoc.GetResourceValue("Description") & "</TH>" & v_th & IPLoc.GetResourceValue("Initiator") & "</TH>" & v_th & IPLoc.GetResourceValue("Implementation Date") & "</TH>")
                    sbEmailBody.Append("</B></TR>")
                End If

                sbEmailBody.Append("<BR><TR valign=top>")
                sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}{0}" & strDesc)
                'sbEmailBody.Append("{0}" & strDesc & "{1}")
                sbEmailBody.Append("{0}" & strInitiator & "{1}")
                sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                If strRecType = "Pending Your Review" Then
                    sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & "&UserName=" & strUserid & ">" & IPLoc.GetResourceValue("Click the blue link to MOC to acknowledge that you have reviewed the MOC") & "</A>{1}")
                End If

                previous_recType = strRecType
                i = i + 1
            End While

            ' strHeading = "<BR><B>" & IPLoc.GetResourceValue("Draft MOCs") & "</B>"
            If drMOCs.NextResult Then

                While drMOCs.Read
                    strRecType = drMOCs("approval_Type")
                    strMOCID = drMOCs("MOCNumber")
                    strTitle = drMOCs("title")
                    If drMOCs("Description") Is DBNull.Value Then
                        strDesc = ""
                    Else
                        strDesc = drMOCs("Description")
                        strDesc = Replace(strDesc, "{", "(")
                        strDesc = Replace(strDesc, "}", ")")
                    End If
                    If drMOCs("EventDate") Is DBNull.Value Then
                        strImplementationDate = String.Empty
                    Else
                        strImplementationDate = drMOCs("EventDate")
                        strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
                    End If
                    strInitiator = drMOCs("initiatorname")
                    'strEmailAddress = drMOCs("Email")

                    If previous_recType <> strRecType Or previous_recType = "" Then
                        sbEmailBody.Append("</table><P style='    padding: 10px; background-color: #e0e0eb;    border: solid 1px #DDEEEE;    color: #336B6B;'><font size=2 face=Arial><B><U>" & IPLoc.GetResourceValue(strRecType) & "</B></U></FONT>")
                        sbEmailBody.Append(v_table & "<thead><TR><B>" & v_th & IPLoc.GetResourceValue("Title") & "</TH>" & v_th & IPLoc.GetResourceValue("Description") & "</TH>" & v_th & IPLoc.GetResourceValue("Initiator") & "</TH>" & v_th & IPLoc.GetResourceValue("Implementation Date") & "</TH>")
                        sbEmailBody.Append("</B></TR>")
                    End If

                    sbEmailBody.Append("<BR><TR valign=top>")
                    sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}{0}" & strDesc & "{1}")
                    sbEmailBody.Append("{0}" & strInitiator & "{1}")
                    sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                    previous_recType = strRecType
                    i = i + 1

                End While
            End If

            ' strHeading = "<BR><B>" & IPLoc.GetResourceValue("On Hold MOCs") & "</B>"
            If drMOCs.NextResult Then

                While drMOCs.Read
                    strRecType = drMOCs("approval_Type")
                    strMOCID = drMOCs("MOCNumber")
                    strTitle = drMOCs("title")
                    If drMOCs("Description") Is DBNull.Value Then
                        strDesc = ""
                    Else
                        strDesc = drMOCs("Description")
                        strDesc = Replace(strDesc, "{", "(")
                        strDesc = Replace(strDesc, "}", ")")
                    End If
                    If drMOCs("EventDate") Is DBNull.Value Then
                        strImplementationDate = String.Empty
                    Else
                        strImplementationDate = drMOCs("EventDate")
                        strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
                    End If
                    strInitiator = drMOCs("initiatorname")
                    'strEmailAddress = drMOCs("Email")

                    If previous_recType <> strRecType Or previous_recType = "" Then
                        sbEmailBody.Append("</table><P style='    padding: 10px; background-color: #e0e0eb;    border: solid 1px #DDEEEE;    color: #336B6B;'><font size=2 face=Arial><B><U>" & IPLoc.GetResourceValue(strRecType) & "</B></U></FONT>")
                        sbEmailBody.Append(v_table & "<thead><TR><B>" & v_th & IPLoc.GetResourceValue("Title") & "</TH>" & v_th & IPLoc.GetResourceValue("Description") & "</TH>" & v_th & IPLoc.GetResourceValue("Initiator") & "</TH>" & v_th & IPLoc.GetResourceValue("Implementation Date") & "</TH>")
                        sbEmailBody.Append("</B></TR>")
                    End If

                    sbEmailBody.Append("<BR><TR valign=top>")
                    sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}{0}" & strDesc & "{1}")
                    sbEmailBody.Append("{0}" & strInitiator & "{1}")
                    sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                    previous_recType = strRecType
                    i = i + 1

                End While
            End If

            ' strHeading = "<BR><B>" & IPLoc.GetResourceValue("Implemented without All Approvals") & "</B>"
            If drMOCs.NextResult Then

                While drMOCs.Read
                    strRecType = drMOCs("approval_Type")
                    strMOCID = drMOCs("MOCNumber")
                    strTitle = drMOCs("title")
                    'strSite = drMOCs("Siteid")
                    If drMOCs("Description") Is DBNull.Value Then
                        strDesc = ""
                    Else
                        strDesc = drMOCs("Description")
                        strDesc = Replace(strDesc, "{", "(")
                        strDesc = Replace(strDesc, "}", ")")
                    End If
                    If drMOCs("EventDate") Is DBNull.Value Then
                        strImplementationDate = String.Empty
                    Else
                        strImplementationDate = drMOCs("EventDate")
                        strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
                    End If
                    strInitiator = drMOCs("initiatorname")
                    'strEmailAddress = drMOCs("Email")

                    If previous_recType <> strRecType Or previous_recType = "" Then
                        sbEmailBody.Append("</table><P style='    padding: 10px; background-color: #e0e0eb;    border: solid 1px #DDEEEE;    color: #336B6B;'><font size=2 face=Arial><B><U>" & IPLoc.GetResourceValue(strRecType) & "</B></U></FONT>")
                        sbEmailBody.Append(v_table & "<thead><TR><B>" & v_th & IPLoc.GetResourceValue("Title") & "</TH>" & v_th & IPLoc.GetResourceValue("Description") & "</TH>" & v_th & IPLoc.GetResourceValue("Initiator") & "</TH>" & v_th & IPLoc.GetResourceValue("Implementation Date") & "</TH>")
                        sbEmailBody.Append("</B></TR>")
                    End If

                    sbEmailBody.Append("<BR><TR valign=top>")
                    sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}{0}" & strDesc & "{1}")
                    sbEmailBody.Append("{0}" & strInitiator & "{1}")
                    sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                    previous_recType = strRecType
                    i = i + 1

                End While
            End If

            '  strHeading = "<BR><B>" & IPLoc.GetResourceValue("Completed without All Approvals") & "</B>"
            If drMOCs.NextResult Then

                While drMOCs.Read()
                    strRecType = drMOCs("approval_Type")
                    strMOCID = drMOCs("MOCNumber")
                    strTitle = drMOCs("title")
                    'strSite = drMOCs("Siteid")
                    If drMOCs("Description") Is DBNull.Value Then
                        strDesc = ""
                    Else
                        strDesc = drMOCs("Description")
                        strDesc = Replace(strDesc, "{", "(")
                        strDesc = Replace(strDesc, "}", ")")
                    End If
                    If drMOCs("EventDate") Is DBNull.Value Then
                        strImplementationDate = String.Empty
                    Else
                        strImplementationDate = drMOCs("EventDate")
                        strImplementationDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strImplementationDate, strLanguage, "dd MMM yyyy")
                    End If
                    strInitiator = drMOCs("initiatorname")
                    'strEmailAddress = drMOCs("Email")

                    If previous_recType <> strRecType Or previous_recType = "" Then
                        sbEmailBody.Append("</table><P style='    padding: 10px; background-color: #e0e0eb;    border: solid 1px #DDEEEE;    color: #336B6B;'><font size=2 face=Arial><B><U>" & IPLoc.GetResourceValue(strRecType) & "</B></U></FONT>")
                        sbEmailBody.Append(v_table & "<thead><TR><B>" & v_th & IPLoc.GetResourceValue("Title") & "</TH>" & v_th & IPLoc.GetResourceValue("Description") & "</TH>" & v_th & IPLoc.GetResourceValue("Initiator") & "</TH>" & v_th & IPLoc.GetResourceValue("Implementation Date") & "</TH>")
                        sbEmailBody.Append("</B></TR>")
                    End If

                    sbEmailBody.Append("<BR><TR valign=top>")
                    sbEmailBody.Append("{0}<A HREF=" & strURL & "/MOC/EnterMOC.aspx?MOCNumber=" & strMOCID & ">" & strMOCID & "-" & strTitle & "</A>{1}{0}" & strDesc & "{1}")
                    sbEmailBody.Append("{0}" & strInitiator & "{1}")
                    sbEmailBody.Append("{0}" & strImplementationDate & "{1}")
                    previous_recType = strRecType
                    i = i + 1

                End While
            End If

            strMsg = sbEmailBody.ToString
            strMsg = String.Format(strMsg, v_td)

            strSubject = IPLoc.GetResourceValue("MOCs - Pending Approval or Review")
            If strDB = "RIDEV" Then
                strHeading1 = "***THIS IS A TEST NOTIFICATION*** MOC BATCH EMAIL RERUN for " & dtRunDate & " " & strEmailAddress
                strEmailAddress = "james.butler@graphicpkg.com"
            End If

            strHeading = "<HTML><head></head><BODY>" & IPLoc.GetResourceValue("The following are MOCs that are pending your approval or review.<BR><BR>Click on Title to update MOC.") & "</B>"
            strHeading = strHeading & "<B> " & IPLoc.GetResourceValue("Please review.") & "</B><br>"
            strFooter = "</HTML></BODY>"

            If dtRunDate <> Now().Date Then
                strBody = strHeading1 & "<P><font size=1 face=Arial><B>MOC BATCH EMAIL RERUN for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR><font size=1>" & strHeading & "<BR>" & strMsg.ToString & strFooter
            Else
                strBody = strHeading1 & "<P><font size=2 face=Arial><B>" & strHeading & "<BR>" & strMsg.ToString & strFooter
            End If

            strBody = cleanString(strBody, "<br>")

            SendEmail(strEmailAddress, "RootCause.FailureAnalysis@graphicpkg.com", strSubject, strBody, "MOC")

            strBody = ""

            drMOCs = Nothing

        Catch ex As Exception
            'Exception handling
            strErr = "Error occurred." & ex.Message
            SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail.MOC", strErr)
            HandleError(strSP, strErr, ex)
        Finally
            connDB.Close()
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub
End Module
