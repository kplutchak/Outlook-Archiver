Imports Microsoft.SqlServer.Management.Smo
Imports Microsoft.SqlServer.Management.Common
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports Microsoft.Exchange.WebServices.Data
Imports Microsoft.Exchange.WebServices.Autodiscover


'Imports Microsoft.Offic

'~~> OUTLINE
'    This code module
'    goes through an Outlook folder, deletes all emails older than a certain date,
'    back up the content of those emails (such as subject, body, from, to, etc.) in
'    an Access database, takes any attachments from those emails and stores them
'    in a local folder, and creates links to the new filepaths of those attachments
'    in the attachment fields in the database

'~~~> CURRENT ISSUES

'~~~> FIXED ISSUES
'     Emails between exchange users no longer have the problem with the FromAddress field.
'     Use SMTP address instead. See code.
'     Fixed issue with small "portrait" thumbnail attachments in messages; images smaller than
'     ?? KB are no longer extracted and saved

'~~~> NOTES
'     After closing Outlook, give Outlook a few seconds to exit before running this code.
'     Otherwise, you may get a "remote server not found" error (may be fixed)
'     The code can only handle 10 attachments per email. If there are more than 10
'     attachments, then the corresponding email is not deleted
'     If Outlook is open when we run, then we keep it open after we are done running
'     Otherwise, we open Outlook to run, then close it on completion
'     The macro stores the deleted emails in the Deleted Items folder in Outlook
'     The user still needs to clear their deleted folder to delete the emails permanentely
'     so they don't take up space.
'     Occasionally,  you may get a "Requesting data from the Exchange server..." or "Outlook is trying
'     to retrieve data…" client popup message in the bottom right hand corner of your screen while the
'     macro is running. This is to be expected - the macro should resume shortly


'~~> REQUIREMENTS
'    TODO

'~~> AUTHORS

'    Kyle Plutchak

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'



Module Module1
    



    Sub Main()
        Dim StartTime As Double, EndTime As Double
        StartTime = Timer

        Try
            Call OutlookToSqlServer()
        Catch ex As Exception
            Exit Sub
        End Try

        EndTime = Timer

        Console.Write("Execution time in seconds: " & EndTime - StartTime & vbCrLf)


    End Sub

    'Check function can be used to verify the integrity of the BLOBs created by this program
    'Sub Check()


    '    Dim conn As New SqlClient.SqlConnection














    '    conn.ConnectionString = "integrated security=SSPI;data source=corp-wsus;persist security info=False;initial catalog=EMAIL-TEST"

    '    conn.Open()

    '    Dim myTrans2 As SqlTransaction
    '    Dim cmd As SqlCommand = conn.CreateCommand

    '    myTrans2 = conn.BeginTransaction()
    '    cmd.Connection = conn
    '    cmd.Transaction = myTrans2

    '    cmd.CommandText = "SELECT * FROM EMAIL_DATA WHERE [EmailID] = 69"
    '    Dim sdr As SqlDataReader = cmd.ExecuteReader()
    '    sdr.Read()



    '    Dim imgByteData As Byte() = CType(sdr.Item("Attachment1"), Byte())
    '    Dim imgMemoryStream As New IO.MemoryStream(imgByteData)
    '    Dim bitmap As System.Drawing.Bitmap = New System.Drawing.Bitmap(imgMemoryStream)

    '    ' or Dim bitmap As Bitmap = Drawing.Image.FromStream(imgMemoryStream)  
    '    bitmap.Save("C:\Users\kplutchak\image.bmp")

    '    sdr.Close()

    '    myTrans2.Commit()

    '    conn.Close()

    'End Sub

    Function strRemoveAddress(ByVal html As String) As String
        Dim returnString As String
        returnString = System.Text.RegularExpressions.Regex.Replace(html, "\<.*?\>", "").Trim()
        Return returnString

    End Function

    Function removeIllegal(ByVal inString As String) As String
        Dim cleanString As String
        cleanString = System.Text.RegularExpressions.Regex.Replace(inString, "[^\w\.@-]", "")
        Return cleanString
    End Function

    Public Function GetPhoto(ByVal filePath As String) As Byte()
        Dim stream As System.IO.FileStream = New System.IO.FileStream( _
           filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim reader As System.IO.BinaryReader = New System.IO.BinaryReader(stream)

        Dim photo() As Byte = reader.ReadBytes(stream.Length)




        reader.Close()
        stream.Close()

        Return photo
    End Function

    Public Function FileFolderExists(ByVal strFilePath As String) As Boolean
        On Error GoTo EarlyExit
        If (Not Dir(strFilePath, vbDirectory) = vbNullString) Then FileFolderExists = True


EarlyExit:
    End Function




    Public Function GetExtension(ByVal cF As String) As String
        Dim cT As String
        Dim tmpString() As String

        cT = StrReverse(cF)
        tmpString = Split(cT, ".")
        GetExtension = "." & StrReverse(tmpString(0))
    End Function

    Private Function CertificateValidationCallBack(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        If sslPolicyErrors = System.Net.Security.SslPolicyErrors.None Then
            Return True
        End If
        If Not sslPolicyErrors = 0 And Not System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors = 0 Then
            If Not chain Is Nothing And Not chain.ChainStatus Is Nothing Then
                For Each status As System.Security.Cryptography.X509Certificates.X509ChainStatus In chain.ChainStatus
                    If certificate.Subject = certificate.Issuer And status.Status = X509ChainStatusFlags.UntrustedRoot Then
                        Continue For
                    ElseIf Not status.Status = System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError Then
                        Return False
                    End If

                Next

            End If
            Return True
        Else
            Return False
        End If
    End Function

    Private Function RedirectionUrlValidationCallback(ByVal redirectionUrl As String) As Boolean
        Dim result As Boolean = False
        Dim redirectionUri As New Uri(redirectionUrl)

        If (redirectionUri.Scheme = "https") Then
            result = True
        End If
        Return result
    End Function



    Public Sub OutlookToSqlServer()


        Dim conn As New SqlClient.SqlConnection

        Dim commandStr As String

        'Debugging variables
        'Dim numAttachmentsSaved As Integer
        'Dim numArchived As Integer


        Dim checkExt As String
        Dim checkSize As Long
        Dim myDeletedFolder As Microsoft.Exchange.WebServices.Data.Folder


        Dim enddate As Date
        Dim startdate As Date 'date to restrict - updated later

        Dim ns As Microsoft.Office.Interop.Outlook.NameSpace
        Dim objFolder As Microsoft.Office.Interop.Outlook.MAPIFolder
        Dim i As Long

        Dim keepOpen As Boolean
        Dim objFolderRestricted As Microsoft.Office.Interop.Outlook.Items






        Dim myemail As Microsoft.Office.Interop.Outlook.MailItem

        Dim temp_path As String
        Dim imageFilepath As String



        Dim ii As Integer
        Dim jj As Integer



        Dim numOther As Integer


        Dim myTrans As SqlTransaction


        Dim objOL As Microsoft.Office.Interop.Outlook.Application



        ' ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack("exchange.nigarawater.com/owa", System.Security.Cryptography.X509Certificates.X509Certificate., )




        Dim service As New ExchangeService(ExchangeVersion.Exchange2010_SP2)

        service.UseDefaultCredentials = True

        service.AutodiscoverUrl("kplutchak@niagarawater.com", AddressOf RedirectionUrlValidationCallback)

        Dim exEmail As EmailMessage

        'Dim userMailbox As New Mailbox("greenteam@niagarawater.com")
        Dim userMailbox As New Mailbox("greenteam@niagarawater.com")

        Dim exFolderID As New Microsoft.Exchange.WebServices.Data.FolderId(WellKnownFolderName.MsgFolderRoot, userMailbox)




        Dim exFolderView As New Microsoft.Exchange.WebServices.Data.FolderView(1000)
        Dim exSubFolderView As New Microsoft.Exchange.WebServices.Data.FolderView(1000)
        'Dim exItemView As New Microsoft.Exchange.WebServices.Data.ItemView(2,  
        Dim exItemView As New Microsoft.Exchange.WebServices.Data.ItemView(1000, 0, OffsetBasePoint.Beginning)

        Dim exSearchFilter As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsEqualTo(Microsoft.Exchange.WebServices.Data.FolderSchema.DisplayName, "Inbox")

        Dim exFolderResults As Microsoft.Exchange.WebServices.Data.FindFoldersResults
        Dim exSubFolderResults As Microsoft.Exchange.WebServices.Data.FindFoldersResults

        Dim myPropertySet As Microsoft.Exchange.WebServices.Data.PropertySet = New PropertySet(BasePropertySet.FirstClassProperties)
        myPropertySet.RequestedBodyType = BodyType.Text

        'enddate = DateAdd(DateInterval.Month, -6, DateTime.Today)
        startdate = CDate(InputBox("Enter a start date"))
        enddate = CDate(InputBox("Enter an end date"))
        'CHANGE

        exFolderResults = service.FindFolders(exFolderID, exSearchFilter, exFolderView)

        exSubFolderResults = exFolderResults(0).FindFolders(exSubFolderView)

        exItemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending)

        Dim searchCollection As Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection = New Microsoft.Exchange.WebServices.Data.SearchFilter.SearchFilterCollection()

        Dim exSearchFilter2 As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsLessThan(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeReceived, enddate)
        Dim exSearchFilter3 As New Microsoft.Exchange.WebServices.Data.SearchFilter.IsGreaterThan(Microsoft.Exchange.WebServices.Data.ItemSchema.DateTimeReceived, startdate)
        Dim mailItems As Microsoft.Exchange.WebServices.Data.FindItemsResults(Of Microsoft.Exchange.WebServices.Data.Item)


        searchCollection.Add(exSearchFilter2)
        searchCollection.Add(exSearchFilter3)




        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        keepOpen = True





        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim todaysdate As String = DateTime.Today.ToString("d")
        todaysdate = todaysdate.Replace("/", "_")
        Dim errorFilePath As String = "C:\Users\kplutchak\archiver_error_file_" & todaysdate & ".txt"
        Dim errorFile As System.IO.FileStream

        Try
            conn.ConnectionString = "integrated security=SSPI;data source=corp-wsus;persist security info=False;initial catalog=EMAIL-TEST"
            conn.Open()
        Catch ex2 As Exception
            Try

                If FileFolderExists(errorFilePath) Then
                    GoTo Line3
                End If
                errorFile = System.IO.File.Create(errorFilePath)
                errorFile.Close()
                My.Computer.FileSystem.WriteAllText(errorFilePath, "Connection with string " & conn.ConnectionString & vbCrLf & ex2.ToString, True)
                GoTo Line3
            Catch ex3 As Exception
                GoTo Line3
            End Try
        End Try







        temp_path = Environ("temp")



        If Not FileFolderExists(temp_path) Then 'CHANGE
            GoTo Line3
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim myCommand As SqlCommand = conn.CreateCommand

        Dim pSubject As New SqlParameter("@sSubject", SqlDbType.VarChar)
        Dim pSender As New SqlParameter("@sSender", SqlDbType.VarChar)
        Dim pTo As New SqlParameter("@sTo", SqlDbType.VarChar)
        Dim pReceived As New SqlParameter("@sReceived", SqlDbType.DateTime)
        Dim pContents As New SqlParameter("@sContents", SqlDbType.Text)
        Dim pCreated As New SqlParameter("@sCreated", SqlDbType.DateTime)
        Dim pModified As New SqlParameter("@sModified", SqlDbType.DateTime)

        Dim pMessageSize As New SqlParameter("@sMessageSize", SqlDbType.Int)
        Dim pFromAddress As New SqlParameter("@sFromAddress", SqlDbType.VarChar)

        Dim pEntryID As New SqlParameter("@sEntryID", SqlDbType.VarChar)
        Dim pCCName As New SqlParameter("@sCCName", SqlDbType.VarChar)


        Dim pNumAttachments As New SqlParameter("@sNumAttachments", SqlDbType.Int)
        Dim pAttachment1 As New SqlParameter("@sAttachment1", SqlDbType.VarBinary)
        Dim pExt1 As New SqlParameter("@sExt1", SqlDbType.VarChar)
        Dim pAttachment2 As New SqlParameter("@sAttachment2", SqlDbType.VarBinary)
        Dim pExt2 As New SqlParameter("@sExt2", SqlDbType.VarChar)
        Dim pAttachment3 As New SqlParameter("@sAttachment3", SqlDbType.VarBinary)
        Dim pExt3 As New SqlParameter("@sExt3", SqlDbType.VarChar)
        Dim pAttachment4 As New SqlParameter("@sAttachment4", SqlDbType.VarBinary)
        Dim pExt4 As New SqlParameter("@sExt4", SqlDbType.VarChar)
        Dim pAttachment5 As New SqlParameter("@sAttachment5", SqlDbType.VarBinary)
        Dim pExt5 As New SqlParameter("@sExt5", SqlDbType.VarChar)
        Dim pAttachment6 As New SqlParameter("@sAttachment6", SqlDbType.VarBinary)
        Dim pExt6 As New SqlParameter("@sExt6", SqlDbType.VarChar)
        Dim pAttachment7 As New SqlParameter("@sAttachment7", SqlDbType.VarBinary)
        Dim pExt7 As New SqlParameter("@sExt7", SqlDbType.VarChar)
        Dim pAttachment8 As New SqlParameter("@sAttachment8", SqlDbType.VarBinary)
        Dim pExt8 As New SqlParameter("@sExt8", SqlDbType.VarChar)
        Dim pAttachment9 As New SqlParameter("@sAttachment9", SqlDbType.VarBinary)
        Dim pExt9 As New SqlParameter("@sExt9", SqlDbType.VarChar)
        Dim pAttachment10 As New SqlParameter("@sAttachment10", SqlDbType.VarBinary)
        Dim pExt10 As New SqlParameter("@sExt10", SqlDbType.VarChar)







        'MsgBox(exFolderResults(0).TotalCount())
        'MsgBox(exFolderResults(0).DisplayName)


        myTrans = conn.BeginTransaction()
        myCommand.Connection = conn
        myCommand.Transaction = myTrans

        Dim exAllFolders As IEnumerable = exSubFolderResults.Concat(exFolderResults)






        For Each myFolder As Microsoft.Exchange.WebServices.Data.Folder In exAllFolders
            Console.WriteLine(myFolder.DisplayName)

        Next

        Dim mailCount As Integer
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For Each myFolder As Microsoft.Exchange.WebServices.Data.Folder In exAllFolders
            'MsgBox(myFolder.DisplayName)
            exItemView.Offset = 0
            If myFolder.TotalCount = 0 Then
                GoTo NextFolder
            End If
            Do
                numOther = 0
                mailItems = myFolder.FindItems(searchCollection, exItemView)

                If mailItems.TotalCount = 0 Then
                    GoTo NextFolder
                End If

                service.LoadPropertiesForItems(mailItems, myPropertySet)

                mailCount = mailItems.Count
                MsgBox(mailCount)
                'For i = mailCount - 1 To 0 Step -1
                For Each myItem As Microsoft.Exchange.WebServices.Data.Item In mailItems

                    'If 1 = 2 Then
                    Console.Write("Archiving email: " & (i + 1).ToString & vbCrLf)
                    If myItem.ItemClass = "IPM.Note" Then


                        exEmail = myItem
                        If exEmail.Attachments.Count < 11 Then

                            Try

                                With exEmail





                                    myCommand.Parameters.Clear()

                                    myCommand.Parameters.Add(pSubject)
                                    myCommand.Parameters.Add(pSender)
                                    myCommand.Parameters.Add(pTo)
                                    myCommand.Parameters.Add(pReceived)
                                    myCommand.Parameters.Add(pContents)
                                    myCommand.Parameters.Add(pCreated)
                                    myCommand.Parameters.Add(pModified)
                                    myCommand.Parameters.Add(pEntryID)

                                    myCommand.Parameters.Add(pMessageSize)
                                    myCommand.Parameters.Add(pFromAddress)

                                    myCommand.Parameters.Add(pCCName)


                                    myCommand.Parameters.Add(pNumAttachments)
                                    If .Attachments.Count > 0 Then
                                        myCommand.Parameters.Add(pAttachment1)
                                        myCommand.Parameters.Add(pExt1)
                                        myCommand.Parameters.Add(pAttachment2)
                                        myCommand.Parameters.Add(pExt2)
                                        myCommand.Parameters.Add(pAttachment3)
                                        myCommand.Parameters.Add(pExt3)
                                        myCommand.Parameters.Add(pAttachment4)
                                        myCommand.Parameters.Add(pExt4)
                                        myCommand.Parameters.Add(pAttachment5)
                                        myCommand.Parameters.Add(pExt5)
                                        myCommand.Parameters.Add(pAttachment6)
                                        myCommand.Parameters.Add(pExt6)
                                        myCommand.Parameters.Add(pAttachment7)
                                        myCommand.Parameters.Add(pExt7)
                                        myCommand.Parameters.Add(pAttachment8)
                                        myCommand.Parameters.Add(pExt8)
                                        myCommand.Parameters.Add(pAttachment9)
                                        myCommand.Parameters.Add(pExt9)
                                        myCommand.Parameters.Add(pAttachment10)
                                        myCommand.Parameters.Add(pExt10)
                                        commandStr = "INSERT INTO EMAIL_DATA ([Subject], [Sender Name], [To], [Received], [Contents], [Created], [Modified],  [Message Size]," & _
                                        " [FromAddress], [EntryID], [CCName], [NumAttachments], [Attachment1]," & _
                                        " [Ext1], [Attachment2], [Ext2], [Attachment3], [Ext3], [Attachment4], [Ext4], [Attachment5], [Ext5], [Attachment6], [Ext6], [Attachment7], [Ext7], [Attachment8], [Ext8], [Attachment9]," & _
                                        " [Ext9], [Attachment10], [Ext10]) VALUES (@sSubject, @sSender, @sTo, @sReceived, @sContents, @sCreated, @sModified," & _
                                        "  @sMessageSize, @sFromAddress, @sEntryID, @sCCName,   @sNumAttachments," & _
                                        " @sAttachment1, @sExt1, @sAttachment2, @sExt2, @sAttachment3, @sExt3, @sAttachment4, @sExt4, @sAttachment5, @sExt5, @sAttachment6, @sExt6," & _
                                        " @sAttachment7, @sExt7, @sAttachment8, @sExt8, @sAttachment9, @sExt9, @sAttachment10, @sExt10)"
                                    Else
                                        commandStr = "INSERT INTO EMAIL_DATA ([Subject], [Sender Name], [To], [Received], [Contents], [Created], [Modified], [Message Size]," & _
                                        " [FromAddress], [EntryID], [CCName],  [NumAttachments])" & _
                                        " VALUES (@sSubject, @sSender, @sTo, @sReceived, @sContents, @sCreated, @sModified," & _
                                        " @sMessageSize, @sFromAddress, @sEntryID, @sCCName, @sNumAttachments)"
                                    End If


                                    myCommand.CommandText = commandStr

                                    If Not .Subject Is Nothing Then
                                        myCommand.Parameters("@sSubject").Value = .Subject
                                    Else
                                        myCommand.Parameters("@sSubject").Value = DBNull.Value
                                    End If

                                    If Not .Sender Is Nothing Then
                                        myCommand.Parameters("@sSender").Value = strRemoveAddress(.From.ToString)
                                    Else
                                        myCommand.Parameters("@sSender").Value = DBNull.Value
                                    End If

                                    If Not .DisplayTo Is Nothing Then
                                        myCommand.Parameters("@sTo").Value = .DisplayTo
                                    Else
                                        myCommand.Parameters("@sTo").Value = DBNull.Value
                                    End If

                                    myCommand.Parameters("@sReceived").Value = .DateTimeReceived

                                    If Not .Body.Text Is Nothing Then

                                        myCommand.Parameters("@sContents").Value = .Body.Text
                                    Else
                                        myCommand.Parameters("@sContents").Value = DBNull.Value
                                    End If

                                    myCommand.Parameters("@sCreated").Value = .DateTimeCreated


                                    myCommand.Parameters("@sModified").Value = .LastModifiedTime





                                    myCommand.Parameters("@sMessageSize").Value = .Size



                                    If .Sender Is Nothing Then
                                        myCommand.Parameters("@sFromAddress").Value = DBNull.Value
                                    Else
                                        myCommand.Parameters("@sFromAddress").Value = .Sender.ToString
                                    End If


                                    myCommand.Parameters("@sEntryID").Value = .Id.ToString


                                    If Not .DisplayCc Is Nothing Then
                                        myCommand.Parameters("@sCCName").Value = .DisplayCc
                                    Else
                                        myCommand.Parameters("@sCCName").Value = DBNull.Value
                                    End If






                                    Dim myFileAttachment As FileAttachment
                                    Dim myItemAttachment As ItemAttachment
                                    Dim myMessage As EmailMessage

                                    If .Attachments.Count > 0 Then
                                        myCommand.Parameters("@sNumAttachments").Value = .Attachments.Count

                                        For ii = 0 To .Attachments.Count - 1




                                            myFileAttachment = TryCast(.Attachments(ii), FileAttachment)
                                            If myFileAttachment Is Nothing Then

                                                myItemAttachment = TryCast(.Attachments(ii), ItemAttachment)
                                                If Not myItemAttachment Is Nothing Then
                                                    myItemAttachment.Load(New PropertySet(ItemSchema.MimeContent))
                                                    'MsgBox(myItemAttachment.GetType.Name)
                                                    'If myItemAttachment.GetType.Name = "EmailMessage" Then
                                                    myMessage = myItemAttachment.Item
                                                    Dim cleanFileName As String
                                                    cleanFileName = temp_path & removeIllegal(myItemAttachment.Name) & ".eml"

                                                    Try
                                                        IO.File.WriteAllBytes(cleanFileName, myMessage.MimeContent.Content)
                                                    Catch itemAttachmentWriteError As Exception
                                                        myCommand.Parameters("@sAttachment" & ii + 1).Value = DBNull.Value
                                                        myCommand.Parameters("@sExt" & ii + 1).Value = DBNull.Value
                                                        GoTo NextAttachment
                                                    End Try

                                                    Dim photo() As Byte = GetPhoto(cleanFileName)
                                                    myCommand.Parameters("@sAttachment" & ii + 1).Value = photo
                                                    My.Computer.FileSystem.DeleteFile(cleanFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                                                    myCommand.Parameters("@sExt" & ii + 1).Value = ".eml"


                                                    GoTo NextAttachment
                                                End If


                                            End If

                                            With myFileAttachment

                                                If .IsContactPhoto() Then
                                                    myCommand.Parameters("@sAttachment" & ii + 1).Value = DBNull.Value
                                                    myCommand.Parameters("@sExt" & ii + 1).Value = DBNull.Value
                                                    GoTo NextAttachment
                                                End If


                                                checkExt = GetExtension(.Name)

                                                checkSize = .Size
                                                imageFilepath = temp_path & .Name

                                                Try
                                                    .Load(temp_path & .Name)

                                                Catch saveFileException As Exception
                                                    myCommand.Parameters("@sAttachment" & ii + 1).Value = DBNull.Value
                                                    myCommand.Parameters("@sExt" & ii + 1).Value = DBNull.Value
                                                    GoTo NextAttachment
                                                End Try

                                                Dim photo() As Byte = GetPhoto(imageFilepath)

                                                myCommand.Parameters("@sAttachment" & ii + 1).Value = photo
                                                My.Computer.FileSystem.DeleteFile(imageFilepath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                                                myCommand.Parameters("@sExt" & ii + 1).Value = checkExt
                                            End With

NextAttachment:

                                        Next ii
                                        For jj = .Attachments.Count + 1 To 10
                                            myCommand.Parameters("@sAttachment" & jj).Value = DBNull.Value
                                            myCommand.Parameters("@sExt" & jj).Value = DBNull.Value

                                        Next jj


                                        'If checkExt = ".jpg" Or checkExt = ".png" Or checkExt = ".gif" Then
                                        '    If checkSize < 10000 Then

                                        '    End If
                                        'End If

                                        'myCommand.Parameters("@sExt1").Value = checkExt
                                        'My.Computer.FileSystem.DeleteFile(imageFilepath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)


                                    Else
                                        myCommand.Parameters("@sNumAttachments").Value = DBNull.Value

                                    End If


                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





                                    myCommand.ExecuteNonQuery()
                                    .Delete(DeleteMode.HardDelete)


                                End With

                            Catch emailException As Exception

                                'Do something
                            End Try



                        End If

                    Else
                        numOther = numOther + 1
                    End If
                    i += 1
                    ' End If 'r
                Next

                If (mailItems.NextPageOffset.HasValue) Then
                    exItemView.Offset = 0
                End If
            Loop Until Not mailItems.MoreAvailable

NextFolder:
        Next

        myTrans.Commit()

        myCommand.Dispose()
        myTrans.Dispose()






        conn.Close()
        conn.Dispose()


        myemail = Nothing





Line3:


        ns = Nothing
        objFolder = Nothing
        objFolderRestricted = Nothing
        objOL = Nothing
        myDeletedFolder = Nothing






    End Sub












End Module


