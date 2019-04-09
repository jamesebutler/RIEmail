
'RIEmail
'
'Application uses EMAILPKG to determine who receives the emails and the 
' content of the emails.  
'
' 5/2010 - Schedule 
' OverdueActions - emails sent every Friday
' OverdueProjects - emails sent every Monday
' Change sub routines to have more robust error handling.
' The Parameters are 
'           strDB = if blank, will connect to production database.  if RIDEV, will connected to dev database.
'           strDefaultEmail = all emails will be sent to this address
'           strEmailBCC = blind carbon copy
'           strSP = stored procedure to call
'           dtProdDate = only used for MOC entered last 7 days
'
' ridev,amy.albrinck@ipaper.com,,OVERDUEPROJECTS,
'
' OVERDUEPROJECTS - Email to Analysis Leader for any incidents that have overdue tasks associated to the incident
' ACTIONSCOMPLETE -
' TEMPMOC - Email to MOC initiator listing any MOCs that are type = Trial/Temporary
' PENDINGMOC
' WEEKLYMOC - Email to Business Unit Areas and approvers for any MOC's that were entered/updated in the
'       past week.  This job was disabled in June 2015 but code left in place in case it needs to be activated.
'
' 10/2015 ALA 
'   Change to use devart dotconnect driver for db calls
'   Removed Overdue Actions references.  Task notifications are all handled in MTTBatchEmail program.
'   Upgrade to .net 4.5.2
'   Consolidated RI/MOC/Outage Batch Emails into this solution.
'
'1/2017 ALA
'   Add Initiator to pending MOC email



Module Main

    Public strDefaultEmail, strDB, strEmailBCC, strSP As String
    Public dtRunDate As Date
    Public strURL As String = "http://gpiri.graphicpkg.com"
    Public strAppName As String = "RIEmail"
    'JEB added 01/31/2019 
    Public supportEmail As String = System.Configuration.ConfigurationManager.AppSettings("supportEmail")
    Public failureEmail As String = System.Configuration.ConfigurationManager.AppSettings("failureEmail")
    Public developmentEmail As String = System.Configuration.ConfigurationManager.AppSettings("developmentEmail")
    Public RIEmail As String = System.Configuration.ConfigurationManager.AppSettings("RIEmail")

    Public tracing As Boolean = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings("Tracing"))







    Sub Main()


        'First Parameter = Database - will default to prod unless RIDEV entered.
        'Second Parameter = default EMAIL  - if you want to run and send all emails to same email account
        'Third Parameter = BCC Email - if you blind cc anyone to check program
        'Forth Parameter = Role
        'Fifth Parameter = RunDate





        Dim cmdLineParams() As String


        '=================================================
        'Used for testing  JEB 01/31/2019
        Dim testit As String = String.Empty
        'testit = ",james.butler@graphicpkg.com,james.butler@graphicpkg.com,OVERDUEPROJECTS,"


        'cmdLineParams = Split(testit, ",")
        '=================================================

        'cmdLineParams = Split(testit, ",")

        Dim i As Integer

        Try
            Dim args As String = ""
            cmdLineParams = Split(Command(), ",")

            For i = 0 To UBound(cmdLineParams)
                If i = 0 Then
                    strDB = cmdLineParams(i)
                    If strDB = "" Then
                        strDB = "RIGPI"
                    Else
                        strDB = "RIDEV"
                    End If
                ElseIf i = 1 Then
                    strDefaultEmail = cmdLineParams(i)
                ElseIf i = 2 Then
                    strEmailBCC = cmdLineParams(i)
                ElseIf i = 3 Then
                    strSP = cmdLineParams(i)
                    strSP = UCase(strSP)
                ElseIf i = 4 Then
                    If cmdLineParams(i) Is DBNull.Value Or cmdLineParams(i) = "" Then
                        dtRunDate = Today()
                    Else
                        dtRunDate = cmdLineParams(i)
                    End If
                End If

            Next


            'Determine which sub routine to go to
            If strSP = "OVERDUEPROJECTS" Then
                'GetOverdueProjects()
            ElseIf strSP = "ACTIONSCOMPLETE" Then
                'GetCertifiedKillActionsComplete()
            ElseIf strSP = "TEMPMOC" Then
                'GetTempMOCUserIds()
            ElseIf strSP = "WEEKLYMOC" Then
                'GetMOCEnteredEmails(dtRunDate)
            ElseIf UCase(strSP) = "PENDINGMOC" Then
                GetPendingMOCUserids()
            End If

            SharedFunctions.SendEmail(developmentEmail, RIEmail, "RIEmail Ran", strSP, "", "", "", True)

        Catch ex As Exception
            HandleError("RIEmail", ex.Message, ex)
        Finally

        End Try

    End Sub

End Module
