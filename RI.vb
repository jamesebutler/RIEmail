Imports System.Net.Mail
Imports System.Configuration
Imports Devart.Common
Imports Devart.Data.Oracle


Module RI


    Sub GetOverdueProjects()
        Dim conCust As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim dr As DataRow
        Dim drEmail As OracleDataReader
        Dim dsEmail As DataSet
        Dim strHeading, strErr As String
        Dim strSubject, strMsg, strBody, strLanguage As String
        Dim strUserid As String = Nothing
        Dim strSiteId As String = Nothing
        Dim strSiteName, strEmailAddress As String
        Dim strEstCompDate, strTaskDesc, strActionLeader As String
        Dim strRinumber, strTitle, strRecordable, strrcfaLevel, strChronic, strSafety, strFooter As String
        Dim previous_rinumber As String
        Dim param As New OracleParameter
        Dim intEmailCnt, intEmailSent As Integer
        Dim sbEmailBody As New System.Text.StringBuilder


        Try
            InsertAuditRecord(strAppName, strSP & " started")

            If strDB = "RIDEV" Then
                conCust.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                conCust.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If

            conCust.Open()

            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = conCust
                .CommandText = "EmailPkg.overdue_projects"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "rsEmail"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            Dim da = New OracleDataAdapter(cmdSQL)
            dsEmail = New DataSet()
            da.Fill(dsEmail)

            'Get email count so we can determine if all emails were sent
            intEmailCnt = dsEmail.Tables(0).Rows.Count()

            'Loop thru records
            For Each dr In dsEmail.Tables(0).Rows

                Try
                    'dr = cmdSQL.ExecuteReader()
                    'While dr.Read
                    sbEmailBody = New System.Text.StringBuilder
                    strLanguage = dr("default_language")
                    If strLanguage = "" Then
                        strLanguage = "EN-US"
                    End If

                    Dim iploc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")

                    strSiteId = dr("siteid")
                    strUserid = dr("project_leader")
                    strSiteName = dr("sitename")
                    strEmailAddress = dr("Email")

                    cmdSQL = New OracleCommand
                    With cmdSQL
                        .Connection = conCust
                        .CommandText = "EmailPkg.BUILD_PROJECT_EMAIL"
                        .CommandType = CommandType.StoredProcedure

                        param = New OracleParameter
                        param.ParameterName = "in_siteid"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strSiteId
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "in_userid"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strUserid
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "email_address"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strEmailAddress
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "in_sitename"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strSiteName
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "rsRecords"
                        param.OracleDbType = OracleDbType.Cursor
                        param.Direction = Data.ParameterDirection.Output
                        .Parameters.Add(param)

                    End With

                    strSubject = iploc.GetResourceValue("Incidents for") & " " & iploc.GetResourceValue(strSiteName) & " " & iploc.GetResourceValue("Require Attention (Analysis Lead)")

                    drEmail = cmdSQL.ExecuteReader()

                    strMsg = ""
                    previous_rinumber = ""
                    While drEmail.Read

                        strSiteId = drEmail("Siteid")
                        strRinumber = drEmail("rinumber")

                        strTitle = drEmail("Incident")
                        strRecordable = drEmail("Recordable")
                        strrcfaLevel = drEmail("RCFALevel")
                        strChronic = drEmail("Chronic")
                        strSafety = drEmail("Safety")
                        strTaskDesc = drEmail("TaskDescription")
                        strEstCompDate = drEmail("EstCompDate")
                        strEstCompDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strEstCompDate, strLanguage)
                        strActionLeader = drEmail("Action_Leader_Name")

                        If previous_rinumber <> strRinumber Then
                            sbEmailBody.Append("<B><U><TD><A HREF=" & strURL & "/RI/ENTERNEWRI.ASPX?RINumber=" & strRinumber & ">" & iploc.GetResourceValue("Incident") & " " & strRinumber & "</A></TD></U></B> ")
                            sbEmailBody.Append("<B>" & iploc.GetResourceValue("IncidentTitle") & ":  </B>" & strTitle & " <BR>")
                            sbEmailBody.Append("<B>" & iploc.GetResourceValue("Recordable") & ": </B>" & iploc.GetResourceValue(strRecordable))
                            sbEmailBody.Append(" <B>" & iploc.GetResourceValue("RCFALevel") & ": </B>" & iploc.GetResourceValue(strrcfaLevel))
                            sbEmailBody.Append(" <B>" & iploc.GetResourceValue("Chronic") & ": </B>" & iploc.GetResourceValue(strChronic))
                            sbEmailBody.Append(" <B>" & iploc.GetResourceValue("EHS") & ": </B>" & iploc.GetResourceValue(strSafety) & "<BR>")
                        End If

                        sbEmailBody.Append("<B>" & iploc.GetResourceValue("Task Description:") & " </B>" & strTaskDesc & "<BR>")
                        sbEmailBody.Append("<B>" & iploc.GetResourceValue("Due Date") & ": </B>" & strEstCompDate)
                        sbEmailBody.Append(" <B>" & iploc.GetResourceValue("Responsible") & ": </B>" & strActionLeader & "<BR><BR>")

                        strMsg = sbEmailBody.ToString
                        'strMsg & strMessage1 & strMessage2 & strMessage3
                        previous_rinumber = strRinumber

                    End While

                    strHeading = "<HTML><BODY><P><B>" & IPLoc.GetResourceValue("OverdueBodyHeading") & "</B><BR>"
                    strFooter = "</HTML></BODY>"
                    strBody = strHeading & "<BR>" & strMsg.ToString & strFooter
                    'Console.WriteLine(strBody & "  ")

                    'strBody = cleanString(strBody, "<br>")
                    SendEmail(strEmailAddress, "RootCause.FailureAnalysis@graphicpkg.com", strSubject, strBody)
                    intEmailSent = intEmailSent + 1


                    'End While

                Catch ex As Exception
                    'Exception handling for any errors that occur when retrieving data and sending email.
                    strErr = "Error occurred on user " & strUserid & "." & ex.Message
                    SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", strErr)
                    HandleError(strSP, strErr, ex)
                End Try

            Next

            'If counts indicate not all emails were sent, send email to support and write record to audit table.
            If intEmailSent <> intEmailCnt Then
                SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "RIBatchEmail", "Only " & intEmailSent & " emails were sent.  " & intEmailCnt & " Emails should have been sent.")
                HandleError(strSP, "Only " & intEmailSent & " emails were sent.  " & intEmailCnt & " Emails should have been sent")
            End If

            dr = Nothing
            drEmail = Nothing
            da = Nothing

            conCust.Close()

            InsertAuditRecord(strAppName, strSP & " ended")

        Catch ex As Exception
            dr = Nothing
            drEmail = Nothing
            HandleError(strSP, ex.Message, ex)
        Finally
            conCust.Close()
            If Not conCust Is Nothing Then conCust = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub
    Sub GetCertifiedKillActionsComplete()
        Dim conCust As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim dr As OracleDataReader = Nothing
        Dim drEmail As OracleDataReader
        Dim strEmailList As String = ""
        Dim strHeading As String = ""
        Dim strSubject As String, strBody As String, strLanguage As String, strErr As String
        Dim strEventDate As String, strActionCompDate As String, strAnalysisCompDate As String
        Dim strRISuperArea As String, strSubArea As String, strArea As String
        Dim strRinumber As String = String.Empty, strTitle As String, strRecordable As String, strrcfaLevel As String, strChronic As String, strCertifiedKill As String, strFooter As String
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder

        Try
            InsertAuditRecord(strAppName, strSP & " started")

            If strDB = "RIDEV" Then
                conCust.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                conCust.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If

            conCust.Open()

            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = conCust
                .CommandText = "EmailPkg.actions_completed"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "rsActionsComp"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            dr = cmdSQL.ExecuteReader()
            While dr.Read
                Try
                    sbEmailBody = New System.Text.StringBuilder
                    strRinumber = dr("rinumber")
                    strLanguage = dr("default_language")
                    If strLanguage = "None" Then
                        strLanguage = "EN-US"
                    End If

                    Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")

                    cmdSQL = New OracleCommand
                    With cmdSQL
                        .Connection = conCust
                        .CommandText = "EmailPkg.BUILD_ACTION_COMP_EMAIL"
                        .CommandType = CommandType.StoredProcedure

                        param = New OracleParameter
                        param.ParameterName = "in_rinumber"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strRinumber
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "in_defaultlanguage"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strLanguage
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "rsRecords"
                        param.OracleDbType = OracleDbType.Cursor
                        param.Direction = Data.ParameterDirection.Output
                        .Parameters.Add(param)

                    End With

                    drEmail = cmdSQL.ExecuteReader()
                    strBody = ""
                    strEmailList = ""

                    While drEmail.Read
                        strEmailList = strEmailList & "," & drEmail("email")
                    End While

                    strRISuperArea = IPLoc.GetResourceValue(dr("risuperarea"))
                    strSubArea = IPLoc.GetResourceValue(dr("subarea"))
                    If Not IsDBNull(dr("Area")) Then
                        strArea = IPLoc.GetResourceValue(dr("Area"))
                    Else
                        strArea = ""
                    End If
                    strEventDate = dr("EventDate")
                    strEventDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strEventDate, strLanguage)

                    strTitle = dr("incident")
                    strRecordable = IPLoc.GetResourceValue(dr("Recordable"))
                    strrcfaLevel = IPLoc.GetResourceValue(dr("RCFALevel"))
                    strChronic = IPLoc.GetResourceValue(dr("Chronic"))
                    strCertifiedKill = IPLoc.GetResourceValue(dr("CertifiedKill"))
                    If Not IsDBNull(dr("RCFAANALYSISCOMPDATE")) Then
                        strAnalysisCompDate = dr("RCFAANALYSISCOMPDATE")
                        strAnalysisCompDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strAnalysisCompDate, strLanguage)
                    Else
                        strAnalysisCompDate = ""
                    End If
                    strActionCompDate = dr("RCFAACTIONCOMPDATE")
                    strActionCompDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strActionCompDate, strLanguage)

                    strSubject = strRinumber & " " & IPLoc.GetResourceValue("Reliability Incident is marked Certified Kill")
                    'strHeading = "<HTML><BODY><P><B>**** THIS IS A TEST INCIDENT NOTIFICATION ****</B>"
                    strHeading = "<HTML><BODY><P><B>" & IPLoc.GetResourceValue("ActionsCompSubHeading") & "</B>"
                    ' ActionsCompSubHeading = All action items have been completed on this Chronic entry and it is deemed Certified Kill.<BR><BR>The Certified Kill flag is marked Yes.

                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("BusinessUnit") & ":</B> " & strRISuperArea & "/" & strSubArea & "/" & strArea)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("IncidentTitle") & ":</B> " & strTitle)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("EventDate") & ":</B> " & strEventDate)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("Recordable") & ":</B> " & strRecordable)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("RCFA") & ":</B> " & strrcfaLevel)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("Chronic") & ":</B> " & strChronic)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("CertifiedKill") & ":</B> " & strCertifiedKill)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("AnalysisCompDateTitle") & ":</B> " & strAnalysisCompDate)
                    sbEmailBody.Append("<BR><B> " & IPLoc.GetResourceValue("ActionsCompDateTitle") & ":</B> " & strActionCompDate)

                    sbEmailBody.Append("<BR><BR><BR><B><U><TD><A HREF=" & strURL & "/RI/EnterNewRI.aspx?RINumber=" & strRinumber & ">" & IPLoc.GetResourceValue("Click here to Review Incident") & "</A></TD></U></B><BR>")

                    strFooter = "</HTML></BODY>"
                    strBody = strHeading & "<BR>" & sbEmailBody.ToString & strFooter
                    '& Mid(strEmailList, 2)
                    strEmailList = Mid(strEmailList, 2)

                    strBody = cleanString(strBody, "<br>")
                    'Console.WriteLine(strBody)

                    If strBody <> "" Then
                        SendEmail(strEmailList, "RootCause.FailureAnalysis@graphicpkg.com", strSubject, strBody)
                        'SendEmail("amy.albrinck@graphicpkg.com", "RootCause.FailureAnalysis@graphicpkg.com", strSubject, strBody)
                    End If

                Catch ex As Exception
                    'Exception handling for any errors that occur when retrieving data and sending email.
                    strErr = "Error occurred " & ex.Message
                    SendEmail(supportEmail, "RootCause.FailureAnalysis@graphicpkg.com", "AutoEmailError", strErr, "MOC")
                    HandleError(strSP, strErr, ex)
                End Try

            End While

            InsertAuditRecord(strAppName, strSP & " ended")

        Catch ex As Exception
            HandleError(strSP, "This attempted email message was not sent b/c :" & ex.Message & "<br>", ex)
        Finally
            If dr IsNot Nothing Then dr = Nothing
        End Try

    End Sub




    'Public Sub HandleError(Optional ByVal MethodName As String = "RI", Optional ByVal additionalErrMsg As String = "", Optional ByVal excep As Exception = Nothing)
    '    Dim le As Exception
    '    Dim errorMessage As New System.Text.StringBuilder
    '    Dim errorCount As Integer = 0
    '    Dim errMsg As String = String.Empty
    '    Dim chunkLength As Integer = 0
    '    Dim maxLen As Integer = 3500
    '    Try
    '        If excep IsNot Nothing Then
    '            le = excep
    '        End If

    '        If le IsNot Nothing Then

    '            Do While le IsNot Nothing
    '                errorCount = errorCount + 1
    '                'errorMessage.Length = 0
    '                errorMessage.Append("<Table width=100% border=1 cellpadding=2 cellspacing=2 bgcolor='#cccccc'>")
    '                errorMessage.Append("<tr><th colspan=2><h2>Error</h2></th>")
    '                errorMessage.Append("<tr><td><b>Program:</b></td><td>{0}</td></tr>")
    '                errorMessage.Append("<tr><td><b>Exception #</b></td><td>{1}</td></tr>")
    '                errorMessage.Append("<tr><td><b>Time:</b></td><td>{2}</td></tr>")
    '                errorMessage.Append("<tr><td><b>Details:</b></td><td>{3}</td></tr>")
    '                errorMessage.Append("<tr><td><b>Additional Info:</b></td><td>{4}</td></tr>")
    '                errorMessage.Append("</table>")
    '                errMsg = errorMessage.ToString
    '                errMsg = String.Format(errMsg, My.Application.Info.AssemblyName, errorCount, FormatDateTime(Now, DateFormat.LongDate), le.ToString, additionalErrMsg)
    '                additionalErrMsg = ""
    '                le = le.InnerException
    '                'MsgBox(errMsg)
    '                errorMessage.Length = 0

    '                For i As Integer = 0 To errMsg.Length Step maxLen
    '                    If errMsg.Length < maxLen Then
    '                        chunkLength = errMsg.Length - 1
    '                    Else
    '                        If errMsg.Length - i < maxLen Then
    '                            chunkLength = errMsg.Length - i
    '                        Else
    '                            chunkLength = maxLen
    '                        End If
    '                    End If

    '                    Dim errValue As String = errMsg.Substring(i, chunkLength)
    '                    'insert record into audit
    '                    InsertAuditRecord(MethodName, errValue)
    '                    System.Threading.Thread.Sleep(1000) ' Sleep for 1 second
    '                Next
    '            Loop
    '        End If

    '    Catch ex As Exception
    '    Finally
    '        le = Nothing
    '        Try
    '            ' HttpContext.Current.Server.ClearError()
    '            'HttpContext.Current.Response.Redirect(redirectURL, False)
    '        Catch e As Exception
    '            'HttpContext.Current.Server.ClearError()
    '        End Try
    '    End Try
    'End Sub


End Module
