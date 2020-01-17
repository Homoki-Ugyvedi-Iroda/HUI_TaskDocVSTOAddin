﻿Imports System.Environment
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Outlook

Module OutlookApplicationHelper
    Public Const Separator = "|"
    Public Const MAPINameSpaceGUID = "http://schemas.microsoft.com/mapi/string/{89465200-caac-429e-931d-6888323b7279}/"
    Public Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

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
    <Extension>
    Friend Function IsNewAndEmpty(ByVal selected As MailItem) As Boolean
        If String.IsNullOrWhiteSpace(selected.Subject) AndAlso String.IsNullOrWhiteSpace(selected.To) Then Return True Else Return False
    End Function
    ''' <summary>
    ''' Először ActiveInspectorban lévő MailItemet küldi vissza (?), ha ez üres, akkor az InlineResponse szerint nyitott emailt, és ha ez is üres, akkor az ActiveExplorerben kijelölt első MailItemet
    ''' </summary>
    ''' <param name="application"></param>
    ''' <returns>A kijelölt MailItem</returns>
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
End Module