' MIT License
'
' Copyright(c) 2019 milCERT
'
' Permission Is hereby granted, free Of charge, to any person obtaining a copy
' of this software And associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, And/Or sell
' copies of the Software, And to permit persons to whom the Software Is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice And this permission notice shall be included In all
' copies Or substantial portions of the Software.
'
' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or
' IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
' LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM,
' OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS IN THE
' SOFTWARE.
<System.Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Microsoft.Office.Core.IRibbonExtensibility

    Private ribbon As Microsoft.Office.Core.IRibbonUI
    Private config As New Config
    Private appLog As New System.Diagnostics.EventLog
    Private keyLanguage As Integer
    Private msgBoxErrorBody As String
    Private msgBoxErrorTitle As String

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Microsoft.Office.Core.IRibbonExtensibility.GetCustomUI
        ' Outlook will try to load all ribbons (found in your ribbon xml) into any window the user goes to. Error if "Show add-in user interface errors" option (in Options -> Advanced).
        Select Case ribbonID
            Case "Microsoft.Outlook.Explorer"
                Return GetResourceText("milcert_spam_add_in.Ribbon.xml")
            Case Else
                Return Nothing
        End Select
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Microsoft.Office.Core.IRibbonUI)
        Me.ribbon = ribbonUI
        keyLanguage = Globals.ThisAddIn.Application.LanguageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)
    End Sub

    Public Function GroupSpam_Label(ByVal control As Microsoft.Office.Core.IRibbonControl) As String
        Return config.group.Item(If(config.group.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
    End Function

    Public Function ButtonSpam_Label(ByVal control As Microsoft.Office.Core.IRibbonControl) As String
        Return config.button.Item(If(config.button.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
    End Function

    Public Function ButtonSpam_Description(ByVal control As Microsoft.Office.Core.IRibbonControl) As String
        Return config.buttonHoverDescription.Item(If(config.buttonHoverDescription.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
    End Function

    Public Function ButtonSpam_ScreenTip(ByVal control As Microsoft.Office.Core.IRibbonControl) As String
        Return config.buttonScreenTip.Item(If(config.buttonScreenTip.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
    End Function

    Public Function ButtonSpam_SuperTip(ByVal control As Microsoft.Office.Core.IRibbonControl) As String
        Return config.buttonSuperTip.Item(If(config.buttonSuperTip.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
    End Function

    Public Sub ButtonSpam_Click(ByVal control As Microsoft.Office.Core.IRibbonControl)
        Try
            appLog.Source = Config.eventLogName

            Dim reportEmailBody As String = ""
            Dim msgBoxItemTypeTitle As String = ""
            Dim msgBoxItemTypeBody As String = ""

            Dim msgBoxConfirmTitle As String = ""
            Dim msgBoxConfirmBodyOne As String = ""
            Dim msgBoxConfirmBodyMore As String = ""

            Dim msgBoxEmptyTitle As String = ""
            Dim msgBoxEmptyBody As String = ""

            Dim msgBoxEncryptedTitle As String = ""
            Dim msgBoxEncryptedBody As String = ""

            Dim msgBoxTooManyRecipients As String = ""
            Dim msbBoxNoInternalMsg As String = ""

            Try
                reportEmailBody = config.reportEmailBody.Item(If(config.reportEmailBody.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))

                msgBoxItemTypeTitle = config.msgBoxItemTypeTitle.Item(If(config.msgBoxItemTypeTitle.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msgBoxItemTypeBody = config.msgBoxItemTypeBody.Item(If(config.msgBoxItemTypeBody.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))

                msgBoxConfirmTitle = config.msgBoxConfirmTitle.Item(If(config.msgBoxConfirmTitle.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msgBoxConfirmBodyOne = config.msgBoxConfirmBodyOne.Item(If(config.msgBoxConfirmBodyOne.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msgBoxConfirmBodyMore = config.msgBoxConfirmBodyMore.Item(If(config.msgBoxConfirmBodyMore.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))

                msgBoxEmptyTitle = config.msgBoxEmptyTitle.Item(If(config.msgBoxEmptyTitle.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msgBoxEmptyBody = config.msgBoxEmptyBody.Item(If(config.msgBoxEmptyBody.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))

                msgBoxEncryptedTitle = config.msgBoxEncryptedTitle.Item(If(config.msgBoxEncryptedTitle.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msgBoxEncryptedBody = config.msgBoxEncryptedBody.Item(If(config.msgBoxEncryptedBody.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))

                msgBoxErrorTitle = config.msgBoxErrorTitle.Item(If(config.msgBoxErrorTitle.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msgBoxErrorBody = config.msgBoxErrorBody.Item(If(config.msgBoxErrorBody.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))

                msgBoxTooManyRecipients = config.msgBoxTooManyRecipients.Item(If(config.msgBoxTooManyRecipients.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
                msbBoxNoInternalMsg = config.msbBoxNoInternalMsg.Item(If(config.msbBoxNoInternalMsg.ContainsKey(keyLanguage), keyLanguage, Config.wdEnglishUS))
            Catch ex As System.Exception
                Try
                    appLog.WriteEntry("Exception while setting language, (could be a KeyNotFoundException) " & ex.Message & ex.StackTrace, System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                Catch appEx As System.Exception

                End Try
            End Try

            Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()

            'Dim ins As Outlook.Inspector = Globals.ThisAddIn.Application.ActiveInspector()

            'If ins IsNot Nothing Then
            '    exp.ClearSelection()
            '    exp.AddToSelection(ins.CurrentItem())
            'End If

            Dim selectionCount = &H0

            ' Avoid an exception if called from the home pane
            Try
                selectionCount = exp.Selection.Count
            Catch ex As System.Exception
                selectionCount = &H0
            End Try

            If selectionCount > &H0 Then
                ' Confirm the submission, no is the default value
                If MsgBox(If(selectionCount > &H1, String.Format(msgBoxConfirmBodyMore, selectionCount), msgBoxConfirmBodyOne), MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton2, msgBoxConfirmTitle) = MsgBoxResult.Yes Then
                    For Each phishEmail As Object In exp.Selection()

                        ' Try to cast the selected item as an Microsoft.Office.Interop.Outlook.MailItem, only emails are supported at this time
                        Try
                            phishEmail = CType(phishEmail, Outlook.MailItem)
                        Catch ex As System.InvalidCastException
                            ' Continue to the next event if the type casting failed
                            MsgBox(msgBoxItemTypeBody, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, msgBoxItemTypeTitle)
                            Try
                                appLog.WriteEntry("Unable to cast item (System.InvalidCastException) " & ex.Message & ex.StackTrace, System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                            Catch appEx As System.Exception

                            End Try
                            Continue For
                        End Try

                        Dim phishEmailSecurityFlags = Config.reportSecurityFlagsNothing
                        ' Try if the email was encrypted, and the certificate was not present to decrypt it
                        Try
                            ' Check the phishing email encrypted and signed flags
                            phishEmailSecurityFlags = CInt(CType(phishEmail, Outlook.MailItem).PropertyAccessor.GetProperty(Config.PR_SECURITY_FLAGS))
#If DEBUG Then
                            System.Diagnostics.Debug.Print("Phishing email signed/encrypted status : " & phishEmailSecurityFlags)
#End If
                        Catch ex As System.Exception
                            MsgBox(msgBoxEncryptedBody, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, msgBoxEncryptedTitle)
                            Try
                                appLog.WriteEntry(msgBoxEncryptedBody, System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                            Catch appEx As System.Exception

                            End Try
                            Continue For
                        End Try

                        ' Phishing email was encrypted, therefore for privacy and confidentiality it cannot be forwarded
                        If CBool(phishEmailSecurityFlags And Config.reportSecurityFlagsEncrypted) Then
                            MsgBox(msgBoxEncryptedBody, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, msgBoxEncryptedTitle)
                            Try
                                appLog.WriteEntry(msgBoxEncryptedBody, System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                            Catch appEx As System.Exception

                            End Try
                            Continue For
                        End If

                        Dim internalMessageOverride As Boolean = False

                        If config.filterInternalMessages Then
                            Try
                                ' If Regex.Match(CType(phishEmail, Outlook.MailItem).SenderEmailAddress, "(^/O=INTERNDOMAIN/OU=EXCHANGE|.*domain.ch)", RegexOptions.IgnoreCase).Success Then
                                ' Configure regex in registry

                                If System.Text.RegularExpressions.Regex.Match(CType(phishEmail, Outlook.MailItem).SenderEmailAddress, config.regexInteralMessages, System.Text.RegularExpressions.RegexOptions.IgnoreCase).Success Then
                                    If MsgBox(String.Format(msbBoxNoInternalMsg, CType(phishEmail, Outlook.MailItem).SenderEmailAddress), MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation Or MsgBoxStyle.DefaultButton2, msgBoxEncryptedTitle) = MsgBoxResult.No Then
                                        Continue For
                                    End If

                                    internalMessageOverride = True

                                    'appLog.WriteEntry(msbBoxNoInternalMsg, System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                                    'Continue For
                                End If
                            Catch ex As System.ArgumentNullException
                                ' SenderEmailAddress is Null, should not be, then let the user report this email as spam
                            End Try
                        End If

                        ' Phishing email was encrypted, ' include it decrypted (encrypt & sign)
                        ' If phishEmailSecurityFlags And Config.reportSecurityFlagsEncrypted Then
                        ' reportEmail.Body += "Phishing email was encrypted"
                        ' reportEmail.Importance = Outlook.OlImportance.olImportanceHigh
                        ' reportEmail.Sensitivity = Outlook.OlSensitivity.olConfidential
                        ' phishEmail.PropertyAccessor.SetProperty(Config.PR_SECURITY_FLAGS, Config.reportSecurityFlagsNothing)
                        ' End If

                        Dim reportEmail As Outlook.MailItem = CType(Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

                        ' Report email general infos
                        reportEmail.Subject = Config.reportEmailSubject
                        reportEmail.To = config.toSecurityTeamCERT
                        reportEmail.CC = config.ccSecurityTeamSpamBit
                        reportEmail.Importance = Outlook.OlImportance.olImportanceLow
                        reportEmail.Sensitivity = Outlook.OlSensitivity.olPersonal
                        reportEmail.Body = reportEmailBody
                        reportEmail.DeleteAfterSubmit = True
                        reportEmail.OriginatorDeliveryReportRequested = False
                        reportEmail.ReadReceiptRequested = False

                        Try
                            ' Phishing email contains too many recipients (>100), and thus cannot be forwarded
                            Dim recipientsCount As Integer = CType(phishEmail, Outlook.MailItem).Recipients.Count
                            If recipientsCount > 100 Then
                                MsgBox(String.Format(msgBoxTooManyRecipients, recipientsCount), MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, msgBoxEncryptedTitle)
                                Try
                                    appLog.WriteEntry(String.Format(msgBoxTooManyRecipients, recipientsCount), System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                                Catch appEx As System.Exception

                                End Try
                                Continue For
                            End If
                        Catch ex As System.Exception
                            ' Recipients count failed
                        End Try

                        ' Save and include the phishing mail
                        CType(phishEmail, Outlook.MailItem).SaveAs(config.spamSavedFilename, Outlook.OlSaveAsType.olMSG)
                        reportEmail.Attachments.Add(config.spamSavedFilename, Outlook.OlAttachmentType.olEmbeddeditem)

                        ' Delete phishing email from temp if exist
                        Try
                            If Not String.IsNullOrEmpty(Dir(config.spamSavedFilename)) Then
                                ' Remove read-only flag if exist
                                SetAttr(config.spamSavedFilename, vbNormal)
                                ' Delete file
                                Kill(config.spamSavedFilename)
#If DEBUG Then
                                System.Diagnostics.Debug.Print(config.spamSavedFilename & " deleted from disk")
#End If
                            End If
                        Catch ex As System.Exception
                            Try
                                appLog.WriteEntry("Unable to delete message from " & config.spamSavedFilename & ex.Message & ex.StackTrace, System.Diagnostics.EventLogEntryType.Warning, Config.eventID)
                            Catch appEx As System.Exception

                            End Try
                        End Try

                        Dim linksCount = 0
                        Try
                            ' Count links in phishing email, could give a few more if the BodyFormat of the reported email is in html
                            linksCount = System.Text.RegularExpressions.Regex.Matches(CType(phishEmail, Outlook.MailItem).Body, "(https?:\/\/|[^https?:\/\/]www\.)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count
                        Catch ex As System.Exception
                            ' Has occured with an empty body
                        End Try

                        If linksCount > 0 Then
                            reportEmail.Body += linksCount & " link" & If(linksCount > 1, "s", "") & " in phishing email"
                            reportEmail.Importance = Outlook.OlImportance.olImportanceNormal
                        End If

#If DEBUG Then
                        System.Diagnostics.Debug.Print("Links count " & linksCount)
#End If

                        ' Count attachments in phishing email
                        Dim attachmentCount = CType(phishEmail, Outlook.MailItem).Attachments.Count
                        If attachmentCount > 0 Then
                            reportEmail.Body += attachmentCount & " attachment" & If(attachmentCount > 1, "s", "") & " in phishing email"
                            reportEmail.Importance = Outlook.OlImportance.olImportanceNormal

                            ' TODO check the datatype (string, listof(string), Set, ...) and maybe change to a new one, or list.Contains() to perform a linear search
                            ' Potentially malicious filename extensions
                            Dim blacklistedExtensions = New String() {".exe", ".pif", ".application", ".gadget", ".msi", ".msp", ".com", ".scr", ".hta", ".cpl", ".msc", ".jar", ".bat", ".cmd", ".vb", ".vbs",
                                ".vbe", ".js", ".jse", ".ws", ".wsf", ".wsc", ".wsh", ".ps1", ".ps1xml", ".ps2", ".ps2xml", ".psc1", ".psc2", ".msh", ".msh1", ".msh2", ".mshxml", ".msh1xml", ".msh2xml", ".scf",
                                ".lnk", ".inf", ".reg",
                                ".doc", ".xls", ".ppt", ".docx", ".xlsx", ".pptx", ".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".pptm", ".potm", ".ppam", ".ppsm", ".sldm", ".pdf",
                                ".htm", ".html", ".xhtml", ".xht", ".mht", ".mhtml", ".maff", ".asp", ".aspx", ".bml", ".cfm", ".cgi", ".ihtml", ".jsp", ".las", ".lasso", ".lassoapp", ".pl", ".php", ".phtml",
                                ".rna", ".r", ".rnx", ".shtml", ".stm",
                                ".iso", ".tar", ".bz2", ".gz", ".lz", ".lzma", ".lzo", ".7z", ".s7z", ".ace", ".afa", ".alz", ".apk", ".arc", ".arj", ".b1", ".ba", ".bh", ".cab", ".car", ".cfs", ".cpt", ".dar", ".dd",
                                ".dgc", ".dmg", ".ear", ".gca", ".ha", ".hki", ".ice", ".jar", ".kgb", ".lzh", ".lha", ".lzx", ".pak", ".partimg", ".paq6", ".paq7", ".paq8", ".pea", ".pim", ".pit", ".qda", ".rar", ".rk",
                                ".sda", ".sea", ".sen", ".sfx", ".shk", ".sit", ".sitx", ".sqx", ".tgz", ".tbz2", ".tlz", ".uca", ".uha", ".war", ".wim", ".xar", ".xp3", ".yz1", ".zip", ".zipx", ".zoo", ".zpaq", ".zz"}

                            ' For each attachment, check if the blocklevel or the extension trigger
                            For Each attachment As Outlook.Attachment In CType(phishEmail, Outlook.MailItem).Attachments
                                Dim attachmentExt = Right(attachment.FileName, Len(attachment.FileName) - InStrRev(attachment.FileName, ".") + &H1)
                                Dim attachmentPrint = attachmentExt

                                ' There is no restriction on the type of the attachment based on its file extension, or there is a restriction on the type of the attachment based on its file extension such that users must first save the attachment to disk before opening it.
                                If CBool(attachment.BlockLevel) Then
                                    reportEmail.Importance = Outlook.OlImportance.olImportanceHigh
                                    attachmentPrint += " Blocklevel"
#If DEBUG Then
                                    System.Diagnostics.Debug.Print("Blocklevel extension " & attachment.FileName)
#End If
                                End If

                                If attachment.FileName.EndsWith("easter.egg", System.StringComparison.InvariantCultureIgnoreCase) Then
                                    MsgBox("Sorry, but Easter eggs are unfortunately too fragile To be transported by email." & vbCrLf & "Thank you For your understanding." & vbCrLf & "Your #milCERT", MsgBoxStyle.Information Or MsgBoxStyle.OkOnly, "Information")
                                    Exit Sub
                                End If

                                ' StringComparaison Enum https://msdn.microsoft.com/en-us/library/system.stringcomparison(v=vs.110).aspx
                                For Each ext In blacklistedExtensions
                                    If attachment.FileName.EndsWith(ext, System.StringComparison.InvariantCultureIgnoreCase) Then
                                        reportEmail.Importance = Outlook.OlImportance.olImportanceHigh
                                        attachmentPrint += " **Blacklisted**"
#If DEBUG Then
                                        System.Diagnostics.Debug.Print("Blacklisted extensions " & attachment.FileName)
#End If
                                        Exit For
                                    End If
                                Next
                                reportEmail.Body += attachmentPrint & " [" & attachment.FileName & "]"
                            Next
                        End If

                        ' Most of the phishing email contains 1-2 links
                        If linksCount > 0 And linksCount < 3 Then
                            reportEmail.Importance = Outlook.OlImportance.olImportanceHigh
                        End If

                        ' Phishing email was signed
                        If CBool(phishEmailSecurityFlags And Config.reportSecurityFlagsSigned) Then
                            reportEmail.Body += "Phishing email was signed"
                            reportEmail.Importance = Outlook.OlImportance.olImportanceHigh
                        End If

                        Dim listUnsubscribe As String = ""
                        Try
                            listUnsubscribe = System.Text.RegularExpressions.Regex.Match(CStr(CType(phishEmail, Outlook.MailItem).PropertyAccessor.GetProperty(Config.PR_TRANSPORT_MESSAGE_HEADERS)), "List-Unsubscribe: (.*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups(1).Value
                            If Not String.IsNullOrEmpty(listUnsubscribe) Then
                                reportEmail.Body += listUnsubscribe
                                reportEmail.Importance = Outlook.OlImportance.olImportanceLow
                            Else
                                ' If Regex.Match(CType(phishEmail, Outlook.MailItem).SenderEmailAddress, "(^/O=INTERNDOMAIN/OU=EXCHANGE|.*domain.ch)", RegexOptions.IgnoreCase).Success Then
                                If System.Text.RegularExpressions.Regex.Match(CType(phishEmail, Outlook.MailItem).SenderEmailAddress, "(\.ch$|\.li$)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Success Then
                                    reportEmail.Body += "Swiss domain used"
                                    reportEmail.Importance = Outlook.OlImportance.olImportanceHigh
                                End If
                            End If
                        Catch ex As System.Exception

                        End Try
#If DEBUG Then
                        ' Attachment count
                        System.Diagnostics.Debug.Print("Attachments count :  " & attachmentCount)

                        ' Print the email importance : 0=low, 1=normal, or 2=high
                        System.Diagnostics.Debug.Print("Email importance " & reportEmail.Importance)
#End If
                        ' Get the spam score from the headers or String.Empty and the reported email priority
                        Dim xSpamImportance As String = config.reportEmailImportance.Item(If(config.reportEmailImportance.ContainsKey(reportEmail.Importance), reportEmail.Importance, &H1))
                        Dim xSpamScore As String = ""

                        Try
                            xSpamScore = System.Text.RegularExpressions.Regex.Match(CStr(CType(phishEmail, Outlook.MailItem).PropertyAccessor.GetProperty(Config.PR_TRANSPORT_MESSAGE_HEADERS)), "X-Spam-Status:.*score=([\d\.-]{1,6})", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups(1).Value
                        Catch ex As System.Exception
                            xSpamScore = ""
                        End Try

                        Try
                            If Not String.IsNullOrEmpty(xSpamScore) Then
                                If CDbl(xSpamScore) >= 5 Then
                                    reportEmail.CC = ""
                                End If
                            End If
                        Catch ex As System.Exception
                            reportEmail.CC = config.ccSecurityTeamSpamBit
                        End Try

                        reportEmail.Subject = Config.reportEmailSubject & " [" & xSpamImportance & "] [" & xSpamScore & "] - '" & CType(phishEmail, Outlook.MailItem).Subject & "'" & If(internalMessageOverride, " [INTERN]", "")

                        ' Include the phishing email headers in the report email body (extract headers)
                        reportEmail.Body += vbCrLf & CStr(CType(phishEmail, Outlook.MailItem).PropertyAccessor.GetProperty(Config.PR_TRANSPORT_MESSAGE_HEADERS))

                        ' Send report email without encrypt and sign (mostly to team mailbox)
                        reportEmail.PropertyAccessor.SetProperty(Config.PR_SECURITY_FLAGS, Config.reportSecurityFlagsNothing)

                        Dim successLogEntry = CType(phishEmail, Outlook.MailItem).Subject & vbCrLf & reportEmail.Body

                        If reportEmail.To IsNot Nothing Or reportEmail.CC IsNot Nothing Then
                            ' Send report and delete spam
                            reportEmail.Send()
                        End If

                        Try
                            CType(phishEmail, Outlook.MailItem).Delete()
                        Catch ex As System.Runtime.InteropServices.COMException
                            ' Occurs in debug mode only ? (Win7 VS2010)
                        End Try
                        Try
                            appLog.WriteEntry("Spam reported successfully " & successLogEntry, System.Diagnostics.EventLogEntryType.Information, Config.eventID)
                        Catch appEx As System.Exception

                        End Try
                    Next
                End If
            Else
                ' No message selected
                MsgBox(msgBoxEmptyBody, MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, msgBoxEmptyTitle)
            End If

        Catch ex As System.Exception
            ' Default exception handler, if an unexpected exception occurs
            MsgBox(msgBoxErrorBody & " " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, msgBoxErrorTitle)

            ' Send the message and stack trace by email only to the RibbonSpamConfig.SecurityTeamEmailMilCERT team-mailbox
#If DEBUG Then
            System.Diagnostics.Debug.Print("Unable to process spam " & ex.Message & ex.StackTrace)
#End If
            Try
                appLog.WriteEntry("Unable to process spam " & ex.Message & ex.StackTrace, System.Diagnostics.EventLogEntryType.Error, Config.eventID)
            Catch appEx As System.Exception

            End Try
            If Not String.IsNullOrEmpty(config.toSecurityTeamCERT) Then
                Dim errorEmail As Outlook.MailItem = CType(Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

                errorEmail.Subject = Config.exceptionEmailSubject & " - 'Exception occurs'"
                errorEmail.To = config.toSecurityTeamCERT
                errorEmail.Body = Environ("USERNAME") & " - " & Environ("COMPUTERNAME") & " - " & Config.targetOS & " - " & Config.addinVersion
                errorEmail.Body += "An Exception occurs " & ex.Message & vbCrLf & ex.StackTrace

                errorEmail.Send()
            End If
        End Try
    End Sub

#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), System.StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As System.IO.StreamReader = New System.IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class