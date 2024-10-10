**Automated Email Distribution via VBA Script using Microsoft Word + Outlook**

---

## 1. Introduction

This project aims to automate the process of sending personalized emails to multiple recipients using the Mail Merge feature in Microsoft Word, coupled with the Outlook application. By utilizing VBA (Visual Basic for Applications), the project streamlines email communication for bulk recipients, enhancing efficiency and accuracy.

## 2. Requirements

- **Software**:
  - Microsoft Word with VBA support.
  - Microsoft Outlook.
- **Data**:
  - An Excel file with a list of recipients, including a column labeled "Email."
- **Knowledge**:
  - Basic understanding of VBA and the Mail Merge functionality.

## 3. Setup Instructions

- **Step 1**: Prepare your Excel file containing recipient details, ensuring the column header for email addresses is labeled "Email."
- **Step 2**: Create a Word document that will serve as the email template, containing the desired email body format.
- **Step 3**: Open the Word document and enable the Developer tab (if not already visible).

## 4. Code Explanation

Below is the VBA code that facilitates the automated email sending process:

```vba
Sub MergeToEmailWithFromAddress()
    Dim strFromAddress As String
    Dim strSubjectLine As String
    Dim olApp As Object ' Late binding for Outlook.Application
    Dim olNamespace As Object ' Late binding for Outlook.Namespace
    Dim olMailItem As Object ' Late binding for Outlook.MailItem
    Dim olAccount As Object ' Late binding for Outlook.Account
    Dim mm As MailMerge
    Dim i As Integer
    Dim dataSource As MailMergeDataSource
    Dim accountList As String
    Dim accountChoice As String
    Dim accounts As Object
    Dim selectedAccount As Object
    Dim emailBody As String
    Dim docContent As String
    
    ' Initialize Outlook Application
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Get a list of available Outlook accounts
    Set accounts = olNamespace.accounts
    accountList = ""
    For Each olAccount In accounts
        accountList = accountList & olAccount.SmtpAddress & vbCrLf
    Next olAccount
    
    ' Prompt user to select the 'From' account from available accounts
    accountChoice = InputBox("Select one of the following accounts for the 'From' address:" & vbCrLf & vbCrLf & accountList, "Select From Account")
    
    ' Validate input for account choice
    If accountChoice = "" Then
        MsgBox "You must select a valid account.", vbExclamation
        Exit Sub
    End If
    
    ' Validate if selected account exists
    Set selectedAccount = Nothing
    For Each olAccount In accounts
        If olAccount.SmtpAddress = accountChoice Then
            Set selectedAccount = olAccount
            Exit For
        End If
    Next olAccount
    
    ' Check if a valid account was selected
    If selectedAccount Is Nothing Then
        MsgBox "The account you entered does not exist. Please try again.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt user for the subject line
    strSubjectLine = InputBox("Enter the subject line for the emails:", "Mail Merge - Subject Line")
    
    ' Validate input for subject line
    If strSubjectLine = "" Then
        MsgBox "Subject line is required.", vbExclamation
        Exit Sub
    End If

    ' Access the mail merge data source
    Set mm = ActiveDocument.MailMerge
    Set dataSource = mm.dataSource
    
    ' Ensure the document is set up as a mail merge
    If mm.MainDocumentType = wdNotAMergeDocument Then
        MsgBox "This document is not set up for mail merge.", vbExclamation
        Exit Sub
    End If
    
    ' Get the full content of the active document for email body
    docContent = ActiveDocument.Content.Text
    
    ' Loop through each record in the mail merge data source
    For i = 1 To dataSource.RecordCount
        dataSource.ActiveRecord = i ' Move to the current record
        
        ' Create a new email
        Set olMailItem = olApp.CreateItem(0) ' 0 refers to Outlook.MailItem type for late binding
        
        ' Set the account to send from the selected account
        Set olMailItem.SendUsingAccount = selectedAccount
        
        ' Set the recipient from the mail merge data source
        olMailItem.To = dataSource.DataFields("Email").Value
        olMailItem.Subject = strSubjectLine
        
        ' Set the dynamic body content without using placeholders
        emailBody = docContent
        
        ' Set the body of the email
        olMailItem.Body = emailBody
        
        ' Send the email
        olMailItem.Send
    Next i
    
    MsgBox "Mail merge completed and emails sent using account: " & accountChoice, vbInformation
End Sub
```

## 5. How to Use the Script

- **Step 1**: Open the Word document containing the VBA script.
- **Step 2**: Ensure that the Mail Merge is set up correctly, pulling recipient data from the Excel file.
- **Step 3**: Run the macro `MergeToEmailWithFromAddress` from the Developer tab.
- **Step 4**: Follow the prompts to select the "From" email account and enter the subject line.
- **Step 5**: The script will send out emails to all recipients listed in the data source.

## 6. Troubleshooting

- **Common Errors**:
  - **Account Not Found**: Ensure that the entered email account matches exactly with one of the available accounts.
  - **Empty Subject Line**: The subject line is required to proceed; ensure it is not left blank.
  - **Not a Mail Merge Document**: Ensure that the active document is set up for mail merge.

## 7. Conclusion

This project successfully demonstrates how to automate email distribution using Microsoft Word and Outlook through the power of VBA. By customizing the email content and leveraging the Mail Merge feature, users can efficiently communicate with multiple recipients.

## 8. References

- [Microsoft VBA Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Outlook Mail Merge Guide](https://support.microsoft.com/en-us/office/use-mail-merge-to-send-personalized-emails-in-word-7f230a98-64d4-4c10-bd8a-b22b1f12b0f4)

---

