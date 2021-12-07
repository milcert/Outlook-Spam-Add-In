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
Imports System.Collections.Generic
Imports Microsoft.Win32

Friend Class Config
    ' Target OS
    Friend Const targetOS As String = "Windows 10 x64"
    ' Current version
    Friend Const addinVersion As String = "1.2.1.0"

    ' SPAM temp saved filename location
    Friend spamSavedFilename As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "spam.msg")
    ' Windows EventLog name
    Friend Const eventLogName = "VSTO 4.0"
    ' Windows EventLod ID
    Friend Const eventID = 1337

    ' Message headers property tag
    Friend Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    ' Message security property tag
    Friend Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"

    ' Registry key that define To and CC addresses
    Friend Const configKey As String = "SOFTWARE\OutlookSpamAddin"
    ' Regex Default
    Private Const regexDefault As String = "(^/O=INTERNDOMAIN/OU=EXCHANGE|(@domain\.ch$|@.*\.domain\.ch$))"

    ' BIT spam email address should be (spam@domain.ch)
    Friend ccSecurityTeamSpamBit As String
    ' milCERT team mailbox
    Friend toSecurityTeamCERT As String
    ' Flag for filtering internal reported emails
    Friend filterInternalMessages As Boolean = True
    ' Regex to filter interal senders
    Friend regexInteralMessages As String
    ' Report email subject
    Friend Const reportEmailSubject As String = "[SPAM]"
    ' Exception email subject
    Friend Const exceptionEmailSubject As String = "[SPAMx]"
    ' Define the max number of recipients allowed to be forwarded (-1 = no limit)
    Friend maxNumberOfRecipients As Integer = -1
    ' If True try to handle encrypted email
    Friend handleEncryptedMailitem As Boolean = False

    ' Report email security flags &H0=nothing, &H1=encrypted, and &H2=signed 
    Friend Const reportSecurityFlagsNothing = &H0
    Friend Const reportSecurityFlagsEncrypted = &H1
    Friend Const reportSecurityFlagsSigned = &H2

    ' User languages https://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
    Friend Const wdGerman As Integer = 1031
    Friend Const wdEnglishUS As Integer = 1033
    Friend Const wdFrench As Integer = 1036
    Friend Const wdItalian As Integer = 1040

    ' Email importance subject
    Friend reportEmailImportance As New Dictionary(Of Integer, String)

    ' Language dictionnaries
    Friend group As New Dictionary(Of Integer, String)
    Friend button As New Dictionary(Of Integer, String)
    Friend buttonHoverDescription As New Dictionary(Of Integer, String)
    Friend buttonScreenTip As New Dictionary(Of Integer, String)
    Friend buttonSuperTip As New Dictionary(Of Integer, String)

    Friend reportEmailBody As New Dictionary(Of Integer, String)

    Friend msgBoxItemTypeBody As New Dictionary(Of Integer, String)
    Friend msgBoxItemTypeTitle As New Dictionary(Of Integer, String)
    Friend msgBoxConfirmTitle As New Dictionary(Of Integer, String)
    Friend msgBoxConfirmBodyOne As New Dictionary(Of Integer, String)
    Friend msgBoxConfirmBodyMore As New Dictionary(Of Integer, String)
    Friend msgBoxEmptyTitle As New Dictionary(Of Integer, String)
    Friend msgBoxEmptyBody As New Dictionary(Of Integer, String)
    Friend msgBoxEncryptedTitle As New Dictionary(Of Integer, String)
    Friend msgBoxEncryptedBody As New Dictionary(Of Integer, String)
    Friend msgBoxErrorTitle As New Dictionary(Of Integer, String)
    Friend msgBoxErrorBody As New Dictionary(Of Integer, String)
    Friend msgBoxTooManyRecipients As New Dictionary(Of Integer, String)
    Friend msbBoxNoInternalMsg As New Dictionary(Of Integer, String)

    Public Sub New()
        Dim regkey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(configKey)

        If regkey IsNot Nothing Then
            toSecurityTeamCERT = CStr(regkey.GetValue("To", ""))
            ccSecurityTeamSpamBit = CStr(regkey.GetValue("Cc", "spam@domain.ch"))
            filterInternalMessages = CBool(regkey.GetValue("FilterInternalMessages", True))
            regexInteralMessages = CStr(regkey.GetValue("Regex", regexDefault))
            maxNumberOfRecipients = CInt(regkey.GetValue("MaxNumberOfRecipients", -1))
            handleEncryptedMailitem = CBool(regkey.GetValue("HandleEncryptedMailItem", False))
        Else
            toSecurityTeamCERT = ""
            ccSecurityTeamSpamBit = "spam@domain.ch"
            filterInternalMessages = True
            regexInteralMessages = regexDefault
            maxNumberOfRecipients = -1
            handleEncryptedMailitem = False
        End If

        msgBoxConfirmBodyOneEN = "Thank you for your contribution to Cyber Security! The selected message will be forwarded to " & ccSecurityTeamSpamBit & " and irrevocably removed from your inbox. Our specialists will take care of it immediately. In case of particularly harmful or dangerous content, you will be contacted personally. Would you like to continue?"
        msgBoxConfirmBodyMoreEN = "Thank you for your contribution to Cyber Security! The selected {0} message will be forwarded to " & ccSecurityTeamSpamBit & " and irrevocably removed from your inbox. Our specialists will take care of it immediately. In case of particularly harmful or dangerous content, you will be contacted personally. Would you like to continue?"
        msgBoxErrorBodyEN = "An error occurred! Please contact " & toSecurityTeamCERT & " to resolve the issue. The selected message was not forwarded nor deleted from your inbox."

        msgBoxConfirmBodyOneDE = "Besten Dank für Ihren Beitrag zur Cyber-Sicherheit! Die ausgewählte Nachricht wird an " & ccSecurityTeamSpamBit & " weitergeleitet und von Ihrem Posteingang unwiderruflich entfernt. Unsere Spezialisten werden sich umgehend darum kümmern. Bei besonders schädlichem oder gefährlichem Inhalt werden Sie persönlich kontaktiert. Möchten Sie weiterfahren?"
        msgBoxConfirmBodyMoreDE = "Besten Dank für Ihren Beitrag zur Cyber-Sicherheit! Die ausgewählte {0} Nachricht wird an " & ccSecurityTeamSpamBit & " weitergeleitet und von Ihrem Posteingang unwiderruflich entfernt. Unsere Spezialisten werden sich umgehend darum kümmern. Bei besonders schädlichem oder gefährlichem Inhalt werden Sie persönlich kontaktiert. Möchten Sie weiterfahren?"
        msgBoxErrorBodyDE = "Ein Fehler ist aufgetreten! Bitte an " & toSecurityTeamCERT & " melden um den Fehler zu beheben. Die ausgewählte Nachricht wurde weder weitergeleited noch vom Posteingang gelöscht."

        msgBoxConfirmBodyOneFR = "Merci pour votre contribution à la cybersécurité! Le message sélectionné sera transmis à " & ccSecurityTeamSpamBit & " et sera irrévocablement supprimé de votre boîte de réception. Nos spécialistes s'en occupent immédiatement. En cas de contenu particulièrement nuisible ou dangereux, vous serez contacté personnellement. Voulez-vous continuer ?"
        msgBoxConfirmBodyMoreFR = "Merci pour votre contribution à la cybersécurité! Le {0} message sélectionné sera transmis à " & ccSecurityTeamSpamBit & " et sera irrévocablement supprimé de votre boîte de réception. Nos spécialistes s'en occupent immédiatement. En cas de contenu particulièrement nuisible ou dangereux, vous serez contacté personnellement. Voulez-vous continuer ?"
        msgBoxErrorBodyFR = "Une erreur est survenue! Veuillez contacter " & toSecurityTeamCERT & " pour résoudre le problème. Le message sélectionné n'a pas été transmis ni supprimé de votre boîte de réception."

        msgBoxConfirmBodyOneIT = "Grazie mille per il suo contributo alla cibersicurezza! L'email selezionata sarà irrevocabilmente eliminata dall'inbox e quindi inoltrata a " & ccSecurityTeamSpamBit & ". I nostri specialisti se ne occuperanno immediatamente. Nel caso di un contenuto particolarmente dannoso o pericoloso sarà contattato personalmente. Vuole procedere?"
        msgBoxConfirmBodyMoreIT = "Grazie mille per il suo contributo alla cibersicurezza! Le {0} email selezionate saranno irrevocabilmente eliminate dall'inbox e quindi inoltrate a " & ccSecurityTeamSpamBit & ". I nostri specialisti se ne occuperanno immediatamente. Nel caso di un contenuto particolarmente dannoso o pericoloso sarà contattato personalmente. Vuole procedere?"
        msgBoxErrorBodyIT = "Si è verificato un errore! Per favore contatti " & toSecurityTeamCERT & " per annunciare il problema. L'email selezionata non è stata né trasmessa, né rimossa dall'inbox."

        If String.IsNullOrEmpty(toSecurityTeamCERT) Then
            msgBoxErrorBodyEN = "An error occured. The selected message was not forwarded nor deleted from your inbox."
            msgBoxErrorBodyDE = "Ein Fehler ist aufgetreten. Die ausgewählte Nachricht wurde weder weitergeleited noch vom Posteingang gelöscht."
            msgBoxErrorBodyFR = "Une erreur est survenue. Le message sélectionné n'a pas été transmis ni supprimé de votre boîte de réception."
            msgBoxErrorBodyIT = "Si è verificato un errore. Lo Spam selezionato non é stato ne trasmesso ne rimosso della sua inbox."
        End If

        If String.IsNullOrEmpty(ccSecurityTeamSpamBit) Then
            msgBoxConfirmBodyOneEN = "The selected message will be forwarded to " & toSecurityTeamCERT & " and removed from your inbox. Would you like to continue?"
            msgBoxConfirmBodyMoreEN = "The {0} selected messages will be forwarded to " & toSecurityTeamCERT & " and removed from your inbox. Would you like to continue?"

            msgBoxConfirmBodyOneDE = "Die ausgewählte Nachricht wird an " & toSecurityTeamCERT & " weitergeleitet und vom Posteingang gelöscht. Weiterfahren?"
            msgBoxConfirmBodyMoreDE = "Die {0} ausgewählten Nachrichten werden an " & toSecurityTeamCERT & " weitergeleitet und vom Posteingang gelöscht. Weiterfahren?"

            msgBoxConfirmBodyOneFR = "Le message sélectionné sera transmis à " & toSecurityTeamCERT & " et supprimé de votre boîte de réception. Voulez-vous continuer ?"
            msgBoxConfirmBodyMoreFR = "Les {0} messages sélectionnés seront transmis à " & toSecurityTeamCERT & " et supprimés de votre boîte de réception. Voulez-vous continuer ?"

            msgBoxConfirmBodyOneIT = "L'email selezionata sarà trasmessa a " & toSecurityTeamCERT & " e rimossa della sua inbox. Vuole procedere?"
            msgBoxConfirmBodyMoreIT = "Le {0} email selezionate sarrano trasmesse a " & toSecurityTeamCERT & " e rimosse della sua inbox. Vuole procedere?"
        End If

        With reportEmailImportance
            .Add(&H0, "L")
            .Add(&H1, "N")
            .Add(&H2, "H")
        End With
        With group
            .Add(wdGerman, groupDE)
            .Add(wdEnglishUS, groupEN)
            .Add(wdFrench, groupFR)
            .Add(wdItalian, groupIT)
        End With
        With button
            .Add(wdGerman, buttonDE)
            .Add(wdEnglishUS, buttonEN)
            .Add(wdFrench, buttonFR)
            .Add(wdItalian, buttonIT)
        End With
        With buttonHoverDescription
            .Add(wdGerman, buttonHoverDescriptionDE)
            .Add(wdEnglishUS, buttonHoverDescriptionEN)
            .Add(wdFrench, buttonHoverDescriptionFR)
            .Add(wdItalian, buttonHoverDescriptionIT)
        End With
        With buttonScreenTip
            .Add(wdGerman, buttonScreenTipDE)
            .Add(wdEnglishUS, buttonScreenTipEN)
            .Add(wdFrench, buttonScreenTipFR)
            .Add(wdItalian, buttonScreenTipIT)
        End With
        With buttonSuperTip
            .Add(wdGerman, buttonSuperTipDE)
            .Add(wdEnglishUS, buttonSuperTipEN)
            .Add(wdFrench, buttonSuperTipFR)
            .Add(wdItalian, buttonSuperTipIT)
        End With

        With reportEmailBody
            .Add(wdGerman, reportEmailBodyDE)
            .Add(wdEnglishUS, reportEmailBodyEN)
            .Add(wdFrench, reportEmailBodyFR)
            .Add(wdItalian, reportEmailBodyIT)
        End With

        With msgBoxItemTypeTitle
            .Add(wdGerman, msgBoxItemTypeTitleDE)
            .Add(wdEnglishUS, msgBoxItemTypeTitleEN)
            .Add(wdFrench, msgBoxItemTypeTitleFR)
            .Add(wdItalian, msgBoxItemTypeTitleIT)
        End With
        With msgBoxItemTypeBody
            .Add(wdGerman, msgBoxItemTypeBodyDE)
            .Add(wdEnglishUS, msgBoxItemTypeBodyEN)
            .Add(wdFrench, msgBoxItemTypeBodyFR)
            .Add(wdItalian, msgBoxItemTypeBodyIT)
        End With
        With msgBoxConfirmTitle
            .Add(wdGerman, msgBoxConfirmTitleDE)
            .Add(wdEnglishUS, msgBoxConfirmTitleEN)
            .Add(wdFrench, msgBoxConfirmTitleFR)
            .Add(wdItalian, msgBoxConfirmTitleIT)
        End With
        With msgBoxConfirmBodyOne
            .Add(wdGerman, msgBoxConfirmBodyOneDE)
            .Add(wdEnglishUS, msgBoxConfirmBodyOneEN)
            .Add(wdFrench, msgBoxConfirmBodyOneFR)
            .Add(wdItalian, msgBoxConfirmBodyOneIT)
        End With
        With msgBoxConfirmBodyMore
            .Add(wdGerman, msgBoxConfirmBodyMoreDE)
            .Add(wdEnglishUS, msgBoxConfirmBodyMoreEN)
            .Add(wdFrench, msgBoxConfirmBodyMoreFR)
            .Add(wdItalian, msgBoxConfirmBodyMoreIT)
        End With
        With msgBoxEmptyTitle
            .Add(wdGerman, msgBoxEmptyTitleDE)
            .Add(wdEnglishUS, msgBoxEmptyTitleEN)
            .Add(wdFrench, msgBoxEmptyTitleFR)
            .Add(wdItalian, msgBoxEmptyTitleIT)
        End With
        With msgBoxEmptyBody
            .Add(wdGerman, msgBoxEmptyBodyDE)
            .Add(wdEnglishUS, msgBoxEmptyBodyEN)
            .Add(wdFrench, msgBoxEmptyBodyFR)
            .Add(wdItalian, msgBoxEmptyBodyIT)
        End With
        With msgBoxEncryptedTitle
            .Add(wdGerman, msgBoxEncryptedTitleDE)
            .Add(wdEnglishUS, msgBoxEncryptedTitleEN)
            .Add(wdFrench, msgBoxEncryptedTitleFR)
            .Add(wdItalian, msgBoxEncryptedTitleIT)
        End With
        With msgBoxEncryptedBody
            .Add(wdGerman, msgBoxEncryptedBodyDE)
            .Add(wdEnglishUS, msgBoxEncryptedBodyEN)
            .Add(wdFrench, msgBoxEncryptedBodyFR)
            .Add(wdItalian, msgBoxEncryptedBodyIT)
        End With
        With msgBoxErrorTitle
            .Add(wdGerman, msgBoxErrorTitleDE)
            .Add(wdEnglishUS, msgBoxErrorTitleEN)
            .Add(wdFrench, msgBoxErrorTitleFR)
            .Add(wdItalian, msgBoxErrorTitleIT)
        End With
        With msgBoxErrorBody
            .Add(wdGerman, msgBoxErrorBodyDE)
            .Add(wdEnglishUS, msgBoxErrorBodyEN)
            .Add(wdFrench, msgBoxErrorBodyFR)
            .Add(wdItalian, msgBoxErrorBodyIT)
        End With
        With msgBoxTooManyRecipients
            .Add(wdGerman, msgBoxTooManyRecipientsDE)
            .Add(wdEnglishUS, msgBoxTooManyRecipientsEN)
            .Add(wdFrench, msgBoxTooManyRecipientsFR)
            .Add(wdItalian, msgBoxTooManyRecipientsIT)
        End With
        With msbBoxNoInternalMsg
            .Add(wdGerman, msbBoxNoInternalMsgDE)
            .Add(wdEnglishUS, msbBoxNoInternalMsgEN)
            .Add(wdFrench, msbBoxNoInternalMsgFR)
            .Add(wdItalian, msbBoxNoInternalMsgIT)
        End With
    End Sub

    ' Ribbon locales EN (default)
    Private Const groupEN As String = "Report Security Issue"
    Private Const buttonEN As String = "Report Spam"
    Private Const buttonHoverDescriptionEN As String = "Report suspicious emails to the information security team."
    Private Const buttonScreenTipEN As String = "Report Spam"
    Private Const buttonSuperTipEN As String = "Use this button to report suspicious emails to the information security team."
    Private Const msgBoxConfirmTitleEN As String = "Report Spam to your security team"
    Private ReadOnly msgBoxConfirmBodyOneEN As String = "The selected message will be forwarded to " & ccSecurityTeamSpamBit & " and removed from your inbox. Would you like to continue?"
    Private ReadOnly msgBoxConfirmBodyMoreEN As String = "The {0} selected messages will be forwarded to " & ccSecurityTeamSpamBit & " and removed from your inbox. Would you like to continue?"
    Private Const msgBoxItemTypeTitleEN As String = "Not a email item"
    Private Const msgBoxItemTypeBodyEN As String = "Only emails can be forwarded to the security team."
    Private Const msgBoxEmptyTitleEN As String = "No message selected"
    Private Const msgBoxEmptyBodyEN As String = "Please select a message to continue."
    Private Const msgBoxEncryptedTitleEN As String = "Warning"
    Private Const msgBoxEncryptedBodyEN As String = "The selected email Is encrypted, therefore it cannot be forwarded due to privacy and confidentiality reasons."
    Private Const msgBoxErrorTitleEN As String = "Error"
    Private ReadOnly msgBoxErrorBodyEN As String = "An error occured, please contact " & toSecurityTeamCERT & " to resolve the issue. The selected Spam was Not forwarded nor deleted from your inbox."
    Private Const msgBoxTooManyRecipientsEN As String = "The selected email has too many recipients ({0}) and cannot be forwarded to your security team."
    Private Const msbBoxNoInternalMsgEN As String = "This email seems to come from within your organisation (sender: {0}), do you really want to report it as a Spam?"
    Private Const reportEmailBodyEN As String = "This Is a user-submitted report of a suspicious email delivered by the milCERT Outlook Spam Plug-In. Please review the attached email." & vbCrLf

    ' Ribbon locales DE
    Private Const groupDE As String = "Sicherheitsvorfall melden"
    Private Const buttonDE As String = "Spam melden"
    Private Const buttonHoverDescriptionDE As String = "Leitet verdächtige Emails an das Sicherheitsteam weiter."
    Private Const buttonScreenTipDE As String = "Spam melden"
    Private Const buttonSuperTipDE As String = "Diese Schaltfläche leitet verdächtige Emails an das Sicherheitsteam weiter."
    Private Const msgBoxConfirmTitleDE As String = "Spam an das Sicherheitsteam melden"
    Private ReadOnly msgBoxConfirmBodyOneDE As String = "Die ausgewählte Nachricht wird an " & ccSecurityTeamSpamBit & " weitergeleitet und vom Posteingang gelöscht. Weiterfahren?"
    Private ReadOnly msgBoxConfirmBodyMoreDE As String = "Die {0} ausgewählten Nachrichten werden an " & ccSecurityTeamSpamBit & " weitergeleitet und vom Posteingang gelöscht. Weiterfahren?"
    Private Const msgBoxItemTypeTitleDE As String = "Ungültige Auswahl"
    Private Const msgBoxItemTypeBodyDE As String = "Es können ausschliesslich Nachrichten an das Sicherheitsteam weitergeleitet werden."
    Private Const msgBoxEmptyTitleDE As String = "Keine Nachricht ausgewählt"
    Private Const msgBoxEmptyBodyDE As String = "Bitte Nachricht auswählen um weiterzufahren."
    Private Const msgBoxEncryptedTitleDE As String = "Achtung"
    Private Const msgBoxEncryptedBodyDE As String = "Die ausgewählte Nachricht ist verschlüsselt, somit kann sie aus Datenschutz- und Vertraulichkeitsgründen nicht weitergeleitet werden."
    Private Const msgBoxErrorTitleDE As String = "Fehler"
    Private ReadOnly msgBoxErrorBodyDE As String = "Ein Fehler ist aufgetreten, bitte an " & toSecurityTeamCERT & " melden um den Fehler zu beheben. Die ausgewählte Nachricht wurde weder weitergeleited noch vom Posteingang gelöscht."
    Private Const msgBoxTooManyRecipientsDE As String = "Die ausgewählte Nachricht hat zuviele Empfänger ({0}) und kann somit nicht an das Sicherheitsteam weitergeleitet werden."
    Private Const msbBoxNoInternalMsgDE As String = "Diese Nachricht scheint aus Ihrer Organisation zu kommen (Absender: {0}), möchten Sie die Nachricht wirklich als Spam melden?"
    Private Const reportEmailBodyDE As String = reportEmailBodyEN

    ' Ribbon locales FR
    Private Const groupFR As String = "Signaler un incident de sécurité"
    Private Const buttonFR As String = "Rapporter un Spam"
    Private Const buttonHoverDescriptionFR As String = "Signaler des emails suspects à votre équipe de sécurité."
    Private Const buttonScreenTipFR As String = "Rapporter un Spam"
    Private Const buttonSuperTipFR As String = "Utilisez ce bouton pour signaler un email suspect à votre équipe de sécurité."
    Private Const msgBoxConfirmTitleFR As String = "Rapporter un Spam à votre équipe de sécurité"
    Private ReadOnly msgBoxConfirmBodyOneFR As String = "Le message sélectionné sera transmis à " & ccSecurityTeamSpamBit & " et supprimé de votre boîte de réception. Voulez-vous continuer ?"
    Private ReadOnly msgBoxConfirmBodyMoreFR As String = "Les {0} messages sélectionnés seront transmis à " & ccSecurityTeamSpamBit & " et supprimés de votre boîte de réception. Voulez-vous continuer ?"
    Private Const msgBoxItemTypeTitleFR As String = "Sélection incorrecte"
    Private Const msgBoxItemTypeBodyFR As String = "Seul un email peut être transmis à votre équipe de sécurité."
    Private Const msgBoxEmptyTitleFR As String = "Aucun message sélectionné"
    Private Const msgBoxEmptyBodyFR As String = "Veuillez choisir un email pour continuer."
    Private Const msgBoxEncryptedTitleFR As String = "Attention"
    Private Const msgBoxEncryptedBodyFR As String = "Le message sélectionné est chiffré, c'est pourquoi il ne peut être transmis pour des raisons de protection des données et de confidentialité."
    Private Const msgBoxErrorTitleFR As String = "Erreur"
    Private ReadOnly msgBoxErrorBodyFR As String = "Une erreur est survenue, veuillez contacter " & toSecurityTeamCERT & " pour résoudre le problème. Le message sélectionné n'a pas été transmis ni supprimé de votre boîte de réception."
    Private Const msgBoxTooManyRecipientsFR As String = "Le message sélectionné comporte trop de destinataires ({0}) et ne peut être transféré à votre équipe de sécurité."
    Private Const msbBoxNoInternalMsgFR As String = "Ce message semble provenir de votre organisation (expéditeur: {0}), voulez-vous vraiment le reporter en tant que Spam ?"
    Private Const reportEmailBodyFR As String = reportEmailBodyEN

    ' Ribbon locales IT
    Private Const groupIT As String = "Segnala un problema di sicurezza"
    Private Const buttonIT As String = "Segnala Spam"
    Private Const buttonHoverDescriptionIT As String = "Segnala un'email sospetta al team di sicurezza informatica."
    Private Const buttonScreenTipIT As String = "Segnala Spam"
    Private Const buttonSuperTipIT As String = "Utilizza il bottone per segnalare email sospette al team di sicurezza informatica."
    Private Const msgBoxConfirmTitleIT As String = "Segnala un'email sospetta al team di sicurezza informatica"
    Private ReadOnly msgBoxConfirmBodyOneIT As String = "L'email selezionata sarà trasmessa a " & ccSecurityTeamSpamBit & " e rimossa della sua inbox. Vuole procedere?"
    Private ReadOnly msgBoxConfirmBodyMoreIT As String = "Le {0} email selezionate sarrano trasmesse a " & ccSecurityTeamSpamBit & " e rimosse della sua inbox. Vuole procedere?"
    Private Const msgBoxItemTypeTitleIT As String = "L'oggetto selezionato non é un'email"
    Private Const msgBoxItemTypeBodyIT As String = "Possono essere trasmesse al team di sicurezza unicamente delle email."
    Private Const msgBoxEmptyTitleIT As String = "Nessun email selezionata"
    Private Const msgBoxEmptyBodyIT As String = "Seleziona un'email per continuare."
    Private Const msgBoxEncryptedTitleIT As String = "Attenzione"
    Private Const msgBoxEncryptedBodyIT As String = "L'email selezionata é criptata, perciò non puo essere trasmessa per ragioni di privacy e di confidenzialità."
    Private Const msgBoxErrorTitleIT As String = "Errore"
    Private ReadOnly msgBoxErrorBodyIT As String = "Si è verificato un errore, per favore contatti " & toSecurityTeamCERT & " per risolvere il problema. Lo Spam selezionato non é stato ne trasmesso ne rimosso della sua inbox."
    Private Const msgBoxTooManyRecipientsIT As String = "Il messaggio contiene troppi destinatari ({0}) e per questa ragione non può esserre inoltrato al vostro team di sicurezza."
    Private Const msbBoxNoInternalMsgIT As String = "L'origine di questa email é interna all'organizzazione (mittente: {0}), vuole veramente segnalarla come Spam?"
    Private Const reportEmailBodyIT As String = reportEmailBodyEN
End Class