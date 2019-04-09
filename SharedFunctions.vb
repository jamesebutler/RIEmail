Imports System.Net.Mail
Imports System.Configuration

Module SharedFunctions
    Function CallDROraclePackage(ByRef parms As OracleParameterCollection, ByVal packageName As String) As String 'OracleClient.OracleDataReader
        Dim conCust As OracleConnection = Nothing
        Dim cmdSQL As OracleCommand = Nothing
        Dim connection As String = String.Empty
        Dim provider As String = String.Empty
        Dim dr As OracleDataReader = Nothing
        Dim cnConnection As OracleConnection = Nothing
        Dim returnParamName As String = String.Empty
        Dim returnValue As String = String.Empty
        Dim returnParms As New Collection
        Try

            If strDB = "RIDEV" Then
                connection = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connection = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            'connection = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString

            cmdSQL = New OracleCommand

            With cmdSQL
                cnConnection = New OracleConnection(connection)
                cnConnection.Open()
                .Connection = cnConnection
                .CommandText = packageName
                .CommandType = CommandType.StoredProcedure
                Dim sb As New System.Text.StringBuilder
                For i As Integer = 0 To parms.Count - 1
                    If parms.Item(i).Value Is Nothing Then parms.Item(i).Value = DBNull.Value
                    Dim parm As New OracleParameter
                    parm.Direction = parms.Item(i).Direction
                    parm.DbType = parms.Item(i).DbType
                    parm.OracleDbType = parms.Item(i).OracleDbType
                    parm.Size = parms.Item(i).Size
                    If parms.Item(i).Direction = ParameterDirection.Input Or parms.Item(i).Direction = ParameterDirection.InputOutput Then
                        If parms.Item(i).Value IsNot Nothing Then
                            parm.Value = parms.Item(i).Value
                            If parm.Value.ToString = "" Then
                                parm.IsNullable = True
                                parm.Value = System.DBNull.Value
                            End If
                        Else
                            If parm.OracleDbType = OracleDbType.NVarChar Then
                                'parm.Value = DBNull.Value
                                'parm.Size = 2
                            End If
                        End If
                    ElseIf parms.Item(i).Direction = ParameterDirection.Output Then
                        returnParms.Add(parms.Item(i).ParameterName)
                        returnParamName = parms.Item(i).ParameterName
                    End If
                    parm.ParameterName = parms.Item(i).ParameterName
                    .Parameters.Add(parm)
                    If sb.Length > 0 Then sb.Append(",")
                    If parm.OracleDbType = OracleDbType.VarChar Then
                        If parm.Value IsNot Nothing Then
                            sb.Append(parm.ParameterName & "= '" & parm.Value.ToString & "' Type=" & parm.OracleDbType.ToString)
                        Else
                            sb.Append(parm.ParameterName & "= '" & "Null" & "' Type=" & parm.OracleDbType.ToString)
                        End If
                    Else
                        If parm.Value IsNot Nothing Then
                            sb.Append(parm.ParameterName & "= '" & parm.Value.ToString & "' Type=" & parm.OracleDbType.ToString)
                        Else
                            sb.Append(parm.ParameterName)
                        End If
                    End If
                    sb.AppendLine()
                Next
            End With

            cmdSQL.ExecuteNonQuery()

            'Populate the original parms collection with the data from the output parameters
            For i As Integer = 0 To returnParms.Count - 1
                parms.Item(cmdSQL.Parameters(returnParms.Item(i)).ToString).Value = cmdSQL.Parameters(returnParms.Item(i)).Value.ToString
            Next
            '// return the return value if there is one
            If returnParamName.Length > 0 Then
                returnValue = cmdSQL.Parameters(returnParamName).Value.ToString
            Else
                returnValue = CStr(0)
            End If

        Catch ex As Exception
            If returnValue.Length = 0 Then returnValue = "Error Occurred"
            If Not conCust Is Nothing Then conCust = Nothing
            HandleError("CallDROraclePackage", ex.Message, ex)
        Finally
            CallDROraclePackage = returnValue
            If Not dr Is Nothing Then dr = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
            If cnConnection IsNot Nothing Then
                If cnConnection.State = ConnectionState.Open Then cnConnection.Close()
                cnConnection = Nothing
            End If
        End Try
    End Function

    Sub SendEmail(ByVal toaddress As String, ByVal fromAddress As String, ByVal subject As String, ByVal body As String, Optional ByVal displayName As String = "", Optional ByVal carbonCopy As String = "", Optional ByVal blindCarbonCopy As String = "", Optional ByVal IsBodyHtml As Boolean = True)
        Dim mail As System.Net.Mail.MailMessage = New MailMessage '= New MailMessage(New MailAddress(fromAddress, displayName), New MailAddress(toaddress))

        Dim OkToSend As Boolean = False
        Dim inputAddress As New System.Text.StringBuilder

        'MsgBox(toaddress)
        Try
            'Comment following line after test runs
            'subject = subject & toaddress

            If strDefaultEmail <> "" Then
                toaddress = strDefaultEmail
            End If
            'toaddress = "amy.albrinck@ipaper.com"

            If toaddress.Length > 0 Then
                Dim toEmail As String() = Split(toaddress, ",")
                For i As Integer = 0 To toEmail.Length - 1
                    If toEmail(i).Length > 0 Then 'And isEmail(toEmail(i)) Then
                        mail.To.Add(toEmail(i))
                    End If
                Next
                If mail.To.Count > 0 Then OkToSend = True
            End If

            'carbonCopy = "cathy.cox@ipaper.com,amy.albrinck@ipaper.com"
            If carbonCopy.Length > 0 Then
                Dim copyEmail As String() = Split(carbonCopy, ",")
                For i As Integer = 0 To copyEmail.Length - 1
                    If copyEmail(i).Length > 0 Then 'And isEmail(copyEmail(i)) Then
                        mail.CC.Add(copyEmail(i))
                    End If
                Next
                If mail.CC.Count > 0 Then OkToSend = True
            End If

            If strEmailBCC <> "" Then
                blindCarbonCopy = strEmailBCC
            End If
            'blindCarbonCopy = "amy.albrinck@ipaper.com"
            If blindCarbonCopy.Length > 0 Then
                Dim bccEmail As String() = Split(blindCarbonCopy, ",")
                For i As Integer = 0 To bccEmail.Length - 1
                    If bccEmail(i).Length > 0 Then 'And isEmail(bccEmail(i)) Then
                        mail.Bcc.Add(bccEmail(i))
                    End If
                Next
                If mail.Bcc.Count > 0 Then OkToSend = True
            End If

            If displayName Is Nothing Or displayName = "" Then
                displayName = "Reliability Incident"
            End If
            If fromAddress.Trim.Length > 0 Then ' And isEmail(fromAddress) Then
                mail.From = New MailAddress(fromAddress, displayName)
            Else
                mail.From = New MailAddress("RootCauseFailureAnalysis@graphicpkg.com", displayName)
            End If
            mail.Priority = MailPriority.High
            mail.IsBodyHtml = IsBodyHtml

            'Send the email message
            mail.Subject = subject
            mail.Body = body

            If OkToSend = True Then
                SharedFunctions.InsertAuditRecord("RI Email", "The following email has been sent -  " & Mid(body, 1, 3000) & "<br> Recipients:" & Mid(toaddress.ToString, 1, 500))
                Dim emailTryCount As Integer = 0
                Dim emailSuccess As Boolean = False
                Do While emailTryCount < 5 And emailSuccess = False
                    Dim client As SmtpClient = New SmtpClient()
                    Try
                        With client
                            emailTryCount += 1
                            .Host = "gpimail.na.graphicpkg.pri"
                            .Timeout = 1000000
                            .Send(mail)
                            emailSuccess = True
                        End With
                        client.Dispose()
                    Catch ex As SmtpException
                        System.Threading.Thread.Sleep(1000)
                    Finally
                        client = Nothing
                    End Try

                Loop
            End If
            mail.Dispose()

        Catch ex As SmtpException
            HandleError("Send Email", "This attempted email message was not sent b/c :" & ex.Message & "<br>" & body & inputAddress.ToString, ex)
        Catch ex As Exception
            HandleError("Send Email", "This attempted email message was not sent b/c :" & ex.Message & "<br>" & body & inputAddress.ToString, ex)
        Finally
            If mail IsNot Nothing Then mail = Nothing
        End Try
    End Sub

    Public Sub HandleError(Optional ByVal MethodName As String = "RIBatchEmail", Optional ByVal additionalErrMsg As String = "", Optional ByVal excep As Exception = Nothing)
        Dim le As Exception = Nothing
        Dim errorMessage As New System.Text.StringBuilder
        Dim errorCount As Integer = 0
        Dim errMsg As String = String.Empty
        Dim chunkLength As Integer = 0
        Dim maxLen As Integer = 3500
        Try
            If excep IsNot Nothing Then
                le = excep
            End If

            If le IsNot Nothing Then

                Do While le IsNot Nothing
                    errorCount = errorCount + 1
                    'errorMessage.Length = 0
                    errorMessage.Append("<Table width=100% border=1 cellpadding=2 cellspacing=2 bgcolor='#cccccc'>")
                    errorMessage.Append("<tr><th colspan=2><h2>Error</h2></th>")
                    errorMessage.Append("<tr><td><b>Program:</b></td><td>{0}</td></tr>")
                    errorMessage.Append("<tr><td><b>Exception #</b></td><td>{1}</td></tr>")
                    errorMessage.Append("<tr><td><b>Time:</b></td><td>{2}</td></tr>")
                    errorMessage.Append("<tr><td><b>Details:</b></td><td>{3}</td></tr>")
                    errorMessage.Append("<tr><td><b>Additional Info:</b></td><td>{4}</td></tr>")
                    errorMessage.Append("</table>")
                    errMsg = errorMessage.ToString
                    errMsg = String.Format(errMsg, My.Application.Info.AssemblyName, errorCount, FormatDateTime(Now, DateFormat.LongDate), le.ToString, additionalErrMsg)
                    additionalErrMsg = ""
                    le = le.InnerException
                    'MsgBox(errMsg)
                    errorMessage.Length = 0

                    For i As Integer = 0 To errMsg.Length Step maxLen
                        If errMsg.Length < maxLen Then
                            chunkLength = errMsg.Length - 1
                        Else
                            If errMsg.Length - i < maxLen Then
                                chunkLength = errMsg.Length - i
                            Else
                                chunkLength = maxLen
                            End If
                        End If

                        Dim errValue As String = errMsg.Substring(i, chunkLength)
                        InsertAuditRecord(MethodName, errValue)
                        SendEmail(supportEmail, "RIEmail@graphicpkg.com", "RIBatchEmail Error", errValue)
                        System.Threading.Thread.Sleep(1000) ' Sleep for 1 second
                    Next
                Loop
            End If

        Catch ex As Exception
        Finally
            le = Nothing
            Try
            Catch ex As Exception
            End Try
        End Try
    End Sub

    Sub InsertAuditRecord(ByVal sourceName As String, ByVal errorMessage As String)
        'INSERT INTO RCFA_AUDIT_LOG VALUES ('DeleteRINumber', SYSDATE, SUBSTR(V_ERRMSG,1,1000) );
        Dim paramCollection As New OracleParameterCollection
        Dim param As New OracleParameter
        Dim ds As System.Data.DataSet = Nothing

        Try

            param = New OracleParameter
            param.ParameterName = "in_name"
            param.OracleDbType = OracleDbType.VarChar
            param.Direction = Data.ParameterDirection.Input
            param.Value = sourceName
            paramCollection.Add(param)

            param = New OracleParameter
            param.ParameterName = "in_desc"
            param.OracleDbType = OracleDbType.VarChar
            param.Direction = Data.ParameterDirection.Input
            param.Value = errorMessage
            paramCollection.Add(param)

            Dim returnStatus As String = CallDROraclePackage(paramCollection, "Reladmin.RIAUDIT.InsertErrorRecord")
        Catch ex As Exception
        Finally
            param = Nothing
            paramCollection = Nothing
        End Try
    End Sub

    Public Function cleanString(ByVal strEdit As String, ByVal defaultValue As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(strEdit, "[\n]", defaultValue).Trim
    End Function

    Public Function ReadNullAsEmptyString(ByVal reader As OracleDataReader, ByVal fieldName As String) As String
        If IsDBNull(reader(fieldName)) Then
            Return String.Empty
        Else
            Return reader(fieldName)
        End If
        Return False
    End Function


End Module
