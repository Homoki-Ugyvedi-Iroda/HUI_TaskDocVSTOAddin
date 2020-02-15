Imports System.Environment
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Outlook
Imports SPHelper.SPHUI
Imports SPHelper.SPFileFolder

Module OutlookApplicationHelper
    Public Const Separator = "|"
    Public Const MAPINameSpaceGUID = "http://schemas.microsoft.com/mapi/string/{89465200-caac-429e-931d-6888323b7279}/"
    Public Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    Public Class OutlookItemAttachment
        Property ID As Integer
        Property FileName As String
        Property DisplayName As String
        Property Size As Integer
    End Class

    Friend Function ConvertPriorityFromSp(InputChoice As String) As OlImportance
        Dim result As New OlImportance
        If IsNothing(InputChoice) Then Return OlImportance.olImportanceNormal
        Select Case InputChoice
            Case "(2) Normal"
                result = OlImportance.olImportanceNormal
            Case "(1) High"
                result = OlImportance.olImportanceHigh
            Case "(3) Low"
                result = OlImportance.olImportanceLow
        End Select
        Return result
    End Function
    Friend Function GetSignature(Format As OlBodyFormat) As String
        Dim appData As String = GetFolderPath(SpecialFolder.ApplicationData)
        Dim SignaturePath = Path.Combine(appData, "Microsoft", "Signatures")
        Dim FormatinString As String = String.Empty
        Select Case Format
            Case OlBodyFormat.olFormatHTML, OlBodyFormat.olFormatUnspecified
                FormatinString = "*.htm"
            Case OlBodyFormat.olFormatPlain
                FormatinString = "*.txt"
                'Case OlBodyFormat.olFormatRichText
                '   FormatinString = "*.rtf"
        End Select
        Dim SignatureFiles As String() = Directory.GetFiles(SignaturePath, FormatinString)
        If SignatureFiles.Length < 1 Then Return String.Empty
        Dim sr As New StreamReader(SignatureFiles(0), Encoding.[Default])
        Dim Signature = sr.ReadToEnd()
        Return Signature
    End Function
    <Extension>
    Friend Function GetCurrentUserInfoTester(application As Outlook.Application) As String
        Dim addrEntry As Outlook.AddressEntry = application.Session.CurrentUser.AddressEntry
        If addrEntry.Type = "EX" Then
            Dim currentUser As Outlook.ExchangeUser = application.Session.CurrentUser.AddressEntry.GetExchangeUser()
            Dim sb As StringBuilder = New StringBuilder()
            If currentUser IsNot Nothing Then
                sb.AppendLine("Name: " & currentUser.Name)
                sb.AppendLine("*STMP address: " & currentUser.PrimarySmtpAddress)
                sb.AppendLine("Title: " & currentUser.JobTitle)
                sb.AppendLine("Department: " & currentUser.Department)
                sb.AppendLine("Location: " & currentUser.OfficeLocation)
                sb.AppendLine("Business phone: " & currentUser.BusinessTelephoneNumber)
                sb.AppendLine("Mobile phone: " & currentUser.MobileTelephoneNumber)
            End If
            Return sb.ToString()
        End If
        Return String.Empty
    End Function
    <Extension>
    Friend Function DisplayAccountInformation(ByVal application As Outlook.Application) As String
        Dim accounts As Outlook.Accounts = application.Session.Accounts
        ' Concatenate a message with information about all accounts.
        Dim builder As StringBuilder = New StringBuilder()
        ' Loop over all accounts and print detail account information.
        ' All properties of the Account object are read-only.
        Dim account As Outlook.Account
        For Each account In accounts
            ' The DisplayName property represents the friendly name of the account.
            builder.AppendFormat("*DisplayName: {0}" & vbNewLine, account.DisplayName)
            ' The UserName property provides an account-based context to determine identity.
            builder.AppendFormat("UserName: {0}" & vbNewLine, account.UserName)
            ' The SmtpAddress property provides the SMTP address for the account.
            builder.AppendFormat("*SmtpAddress: {0}" & vbNewLine, account.SmtpAddress)
            ' The AccountType property indicates the type of the account.
            builder.Append("AccountType: ")
            Select Case (account.AccountType)
                Case Outlook.OlAccountType.olExchange
                    builder.AppendLine("*Exchange")
                Case Outlook.OlAccountType.olHttp
                    builder.AppendLine("Http")
                Case Outlook.OlAccountType.olImap
                    builder.AppendLine("Imap")
                Case Outlook.OlAccountType.olOtherAccount
                    builder.AppendLine("Other")
                Case Outlook.OlAccountType.olPop3
                    builder.AppendLine("Pop3")
            End Select
            builder.AppendLine()
        Next
        Return builder.ToString()
    End Function
    <Extension>
    Friend Function GetUserName(application As Outlook.Application) As String
        Dim addrEntry As Outlook.AddressEntry = application.Session.CurrentUser.AddressEntry
        If addrEntry.Type = "EX" Then
            Dim currentUser As Outlook.ExchangeUser = addrEntry.GetExchangeUser()
            If currentUser IsNot Nothing Then
                Return currentUser.Name
            End If
        End If
        Return String.Empty
    End Function
    <Extension>
    Friend Function GetUserNameSmtp(application As Outlook.Application) As String
        Dim addrEntry As Outlook.AddressEntry = application.Session.CurrentUser.AddressEntry
        If addrEntry.Type = "EX" Then
            Dim currentUser As Outlook.ExchangeUser = addrEntry.GetExchangeUser()
            If currentUser IsNot Nothing Then
                Return currentUser.PrimarySmtpAddress
            End If
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Visszajelzi, hogy az adott MailItem új tétel-e vagy beérkezett, elküldött stb. email
    ''' </summary>
    ''' <returns>True = új, False = nem új, hanem beérkezett/elküldött email (létező item)</returns>
    <Extension>
    Friend Function IsNewAndEmpty(ByVal selected As MailItem) As Boolean
        If String.IsNullOrWhiteSpace(selected.Subject) AndAlso String.IsNullOrWhiteSpace(selected.To) Then Return True Else Return False
    End Function

    ''' <summary>
    ''' Először ActiveInspectorban lévő MailItemet küldi vissza (?), ha ez üres, akkor az InlineResponse szerint nyitott emailt, és ha ez is üres, akkor az ActiveExplorerben kijelölt első MailItemet
    ''' </summary>
    ''' <param name="application"></param>
    ''' <returns>A kijelölt MailItem</returns>
    ''' Egyszerűsíthető, ha Me.OutlookItem-et küldjük neki inkább?

    <Extension>
    Friend Function GetSelectedMailItem(ByVal application As Outlook.Application) As MailItem
        Dim ActiveInspectorCurrentItem = Nothing
        Dim InspectorMailItem As MailItem = Nothing
        Dim ExplorerMailItem As MailItem = Nothing
        Dim myOlSel As Outlook.Selection = Nothing

        If Not IsNothing(application.ActiveInspector) Then ActiveInspectorCurrentItem = application.ActiveInspector.CurrentItem
        If Not IsNothing(ActiveInspectorCurrentItem) AndAlso TypeOf ActiveInspectorCurrentItem Is MailItem Then InspectorMailItem = CType(ActiveInspectorCurrentItem, MailItem)

        Dim ActiveExplorer As Explorer = application.ActiveExplorer
        If Not IsNothing(ActiveExplorer) Then
            If Not IsNothing(ActiveExplorer.ActiveInlineResponse) AndAlso TypeOf ActiveExplorer.ActiveInlineResponse Is MailItem Then
                ExplorerMailItem = CType(ActiveExplorer.ActiveInlineResponse, MailItem)
            Else
                Dim MyExplorerSelection = ActiveExplorer.Selection
                If Not IsNothing(MyExplorerSelection) AndAlso MyExplorerSelection.Count > 0 Then
                    For Each SelectedItem In MyExplorerSelection
                        If SelectedItem.Class = OlObjectClass.olMail Then
                            ExplorerMailItem = TryCast(SelectedItem, MailItem)
                            If IsNothing(ExplorerMailItem) Then Exit For
                        End If
                    Next
                End If
                Marshal.ReleaseComObject(MyExplorerSelection)
            End If
        End If
        If IsNothing(InspectorMailItem) Then Return ExplorerMailItem Else Return InspectorMailItem
    End Function
#Region "GetResultsfromMailItem"
    Public Sub AddCategoryToMail(mailItem As MailItem, valueToAdd As String)
        If String.IsNullOrWhiteSpace(mailItem.Categories) Then
            mailItem.Categories += valueToAdd
            mailItem.Save()
            'Trace.WriteLine("AddCategoryMail_NoCat_mailItemSave AFTER: " & mailItem.Body)
            Exit Sub
        End If
        If mailItem.Categories.Contains(valueToAdd) Then Exit Sub
        mailItem.Categories += "; " & valueToAdd
        mailItem.Save()
        'Trace.WriteLine("AddCategoryMail_NewCat_mailItemSave AFTER: " & mailItem.Body)
        'Marshal.ReleaseComObject(mailItem)
        'mailItem = Nothing
    End Sub
    Public Function GetAllAddress(sourceMailItem As MailItem, RecipientType As OlMailRecipientType) As List(Of String)
        Dim result As New List(Of String)
        For Each recipient As Recipient In sourceMailItem.Recipients
            If recipient.Type = RecipientType Then result.Add(recipient.Address)
        Next
        Return result
    End Function
    Private Function GetMailBody(_sourceItem As MailItem) As String
        If Not IsNothing(_sourceItem) AndAlso Not IsNothing(_sourceItem.Body) Then
            Return _sourceItem.Body.ToString()
        Else
            Return Nothing
        End If
    End Function
    Public Function GetTaskClassesFromMailBody(_sourceItem As MailItem) As IEnumerable(Of TaskClass)
        Dim MySPHUI = Globals.ThisAddIn.Connection
        Dim MailBody As String = GetMailBody(_sourceItem)
        If String.IsNullOrWhiteSpace(MailBody) Then Return Nothing

        Dim resultStr = MatchesToListofString(Regex.Matches(MailBody, RE.TaskRegExp))
        Dim resultTaskClasses As New List(Of TaskClass)
        For Each idString In resultStr
            Dim singleTask As TaskClass = MySPHUI.ConvertTaskIdToTaskClass(MySPHUI.ConvertTaskIdToInteger(idString))
            If Not IsNothing(singleTask) Then resultTaskClasses.Add(singleTask)
        Next
        Return resultTaskClasses
    End Function
    Public Function GetDefaultNonMatterPartnersfromBodyText(_sourceItem As MailItem, selectedMatter As MatterClass) As IEnumerable(Of PersonClass)
        Dim MySPHUI = Globals.ThisAddIn.Connection
        Dim MailBody As String = GetMailBody(_sourceItem)
        If String.IsNullOrWhiteSpace(MailBody) Then Return Nothing
        If IsNothing(selectedMatter) Then Return Nothing

        Dim result As New List(Of PersonClass)
        Dim ListToSearch As New List(Of PersonClass)
        For Each Partner As PersonClass In MySPHUI.Persons.Where(Function(x) Not x.Matter.Equals(selectedMatter))
            Dim PartnerLength = Partner.Value.Length
            Dim InterimStr = Partner.Value.Replace(" ", "")
            If InterimStr.Length < 3 Then Continue For
            ListToSearch.Add(Partner)
        Next
        For Each listelement In ListToSearch
            If Not IsNothing(listelement.CodeName) AndAlso MailBody.Contains(listelement.CodeName) Then
                result.Add(listelement)
                Continue For
            End If
            If listelement.Value.Length > 6 AndAlso result.Contains(listelement) = False AndAlso MailBody.Contains(listelement.Value) Then
                result.Add(listelement)
            End If
        Next
        Return result
    End Function

    Public Function GetMatterfromSenderData(_MailItem As MailItem) As List(Of MatterClass)
        'Dim _MailItem = GetMailItemSelectedinActiveExplorer()
        If IsNothing(_MailItem) Then Return Nothing
        Dim result As New List(Of MatterClass)
        Dim TextSenderMailAddress = _MailItem.SenderEmailAddress
        Dim TextSenderMailDisplayName = _MailItem.SenderName
        Dim senderregexps = Globals.ThisAddIn.Connection.Matters.Where(Function(x) x.SenderAddress <> String.Empty)
        If String.IsNullOrWhiteSpace(TextSenderMailAddress) AndAlso Not String.IsNullOrWhiteSpace(TextSenderMailDisplayName) Then TextSenderMailAddress = TextSenderMailDisplayName
        For Each matter In senderregexps
            Dim RegexSplitArray = matter.SenderAddress.Split(";")
            For Each RegexIndividual In RegexSplitArray
                Dim regex As Regex = New Regex(RegexIndividual)
                Dim match1 As Match = regex.Match(TextSenderMailAddress)
                If match1.Success Then
                    result.Add(matter)
                    Continue For
                End If
            Next
        Next
        Return result
    End Function

#End Region
#Region "Regexp Related Functions"
    Private Function MatchesToListofString(input As MatchCollection) As List(Of String)
        Dim listofmatches As List(Of Match) = input.Cast(Of Match).ToList
        Return listofmatches.Select(Function(m) m.Value).ToList
    End Function
#End Region
#Region "Handling saving attachments and mail"
    Public Function GetAllAttachmentNames(sourceMailItem As MailItem) As List(Of String)
        Dim result As New List(Of String)
        For Each att As Microsoft.Office.Interop.Outlook.Attachment In sourceMailItem.Attachments
            result.Add(GetAttachmentName(att))
        Next
        Return result
    End Function
   
    Private Function IsThisLogo(attachment As Microsoft.Office.Interop.Outlook.Attachment)
        If attachment.FileName.StartsWith("image", StringComparison.InvariantCultureIgnoreCase) AndAlso attachment.Size < 25 * 1024 AndAlso
                (attachment.FileName.EndsWith(".jpg", StringComparison.InvariantCultureIgnoreCase) Or
                attachment.FileName.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase) Or attachment.FileName.EndsWith(".gif")) Then Return True
        Return False
    End Function
    ''' <summary>
    ''' Retrieves only the file names of all attachments of a MailItem
    ''' 
    ''' </summary>
    ''' <param name="mailItem"></param>
    ''' <returns>A list of Strings containing the filenames of the attachments.</returns>
    Public Function GetFileNamesFromAttachmentAsList(mailItem As MailItem) As List(Of String)
        Dim filenames As New List(Of String)
        Dim MailAttachments = mailItem.Attachments
        For Each fname As Outlook.Attachment In mailItem.Attachments
            filenames.Add(fname.FileName)
        Next
        Return filenames
    End Function
    ''' <summary>
    ''' Returns all those attachments that should be saved (that is, excludes small images typical for signatures).
    ''' </summary>
    ''' <param name="_MailItem">The mail item.</param>
    ''' <returns></returns>
    Public Function GetAllAttachmentDisplayNames(_MailItem As MailItem) As List(Of String)
        If IsNothing(_MailItem) Then Return Nothing
        Dim result As New List(Of String)
        For Each att As Microsoft.Office.Interop.Outlook.Attachment In _MailItem.Attachments
            If Not IsThisLogo(att) Then result.Add(att.DisplayName)
        Next
        Return result
    End Function

    ''' <summary>
    ''' Returns all those attachments that should be saved (that is, excludes small images typical for signatures).
    ''' </summary>
    ''' <param name="_MailItem">The mail item.</param>
    ''' <returns></returns>
    Public Function GetAllAttachments(_MailItem As MailItem) As List(Of OutlookItemAttachment)
        If IsNothing(_MailItem) Then Return Nothing
        Dim result As New List(Of OutlookItemAttachment)
        For Each att As Microsoft.Office.Interop.Outlook.Attachment In _MailItem.Attachments
            If Not IsThisLogo(att) Then result.Add(New OutlookItemAttachment With
                                                   {.FileName = att.FileName, .ID = att.Index, .DisplayName = att.DisplayName, .Size = att.Size})
        Next
        Return result
    End Function
    ''' <summary>
    ''' Retrieves the filename (or if there is no filename, a displayname) for a given attachment in an Outlook Item Attachment
    ''' </summary>
    ''' <param name="item">An Attachment</param>
    ''' <returns>The filename or displayname of a given attachment.</returns>
    Public Function GetAttachmentName(item As Microsoft.Office.Interop.Outlook.Attachment) As String
        GetAttachmentName = item.FileName
        If String.IsNullOrWhiteSpace(GetAttachmentName) Then GetAttachmentName = item.DisplayName
        Return ValidateFileName(GetAttachmentName)
    End Function
    ''' <summary>
    ''' Csak magát az emailt, csatolmányok nélkül mentse le egy ideiglenes könyvtárba, és adja vissza egy útvonalként azt, ahova lementette azt.
    ''' </summary>
    ''' <param name="mailItem">Az email, amit le kell menteni csatolmány nélkül HTML formában (.mht)</param>
    ''' <returns>Az útvonala annak a fájlnak, ahova lementette a MailItem-et mht formában</returns>
    Public Function GetMailWOAttachmentsAsFile(mailItem As MailItem) As String
        If IsNothing(mailItem) Then Return Nothing
        Dim fname = Path.Combine(Globals.ThisAddIn.TempPath, GetValidatedMailNameWOExtensionAsFile(mailItem) & ".mht")
        If IO.File.Exists(fname) Then IO.File.Delete(fname)
        mailItem.SaveAs(fname, OlSaveAsType.olMHTML)
        Marshal.ReleaseComObject(mailItem)
        mailItem = Nothing
        Return fname
    End Function
    ''' <summary>
    ''' A csatolmányokkal együtt az emailt magát is mentse le MSG-ként egy ideiglenes könyvtárba, és adjon egy útvonalat arra, hogy hova mentette.
    ''' </summary>
    ''' <param name="mailItem">Az email, amit le kell menteni csatolmányokkal együtt, msg formában (.msg)</param>
    ''' <returns></returns>
    Public Function GetMailwAttachmentsIncludedAsFile(mailItem As MailItem) As String
        If IsNothing(mailItem) Then Return Nothing
        Dim fname = Path.Combine(Globals.ThisAddIn.TempPath) & GetValidatedMailNameWOExtensionAsFile(mailItem) & ".msg"
        mailItem.SaveAs(fname, OlSaveAsType.olMSG)
        Marshal.ReleaseComObject(mailItem)
        mailItem = Nothing
        Return fname
    End Function
    ''' <summary>
    ''' Gets the mail attachments without the mail itself as string of files saved into a temporary directory.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetMailAttachmentWOMailAsFile(mailItem As MailItem, indexOfAttachmentToSave As Integer) As String
        If IsNothing(mailItem) Then Return Nothing
        Dim MailAttachments = mailItem.Attachments
        Dim ListofAttachmentFullPathinTemp As New List(Of String)
        Dim fname = Path.Combine(Globals.ThisAddIn.TempPath, GetAttachmentName(MailAttachments(indexOfAttachmentToSave)))
        If IO.File.Exists(fname) Then IO.File.Delete(fname)
        MailAttachments(indexOfAttachmentToSave).SaveAsFile(fname)
        Marshal.ReleaseComObject(MailAttachments)
        MailAttachments = Nothing
        Return fname
    End Function
    Public Function GetValidatedMailNameWOExtensionAsFile(_MailItem As MailItem) As String
        If IsNothing(_MailItem) Then Return Nothing
        Return ValidateFileName(_MailItem.Subject & "_" & HPHelper.StringRelated.RandomCharGet(4))
    End Function
#End Region

End Module
