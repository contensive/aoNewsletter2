
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterPageClass
        Inherits AddonBaseClass
        '
        ' - update references to your installed version of cpBase
        ' - Verify project root name space is empty
        ' - Change the namespace to the collection name
        ' - Change this class name to the addon name
        ' - Create a Contensive Addon record with the namespace apCollectionName.ad
        ' - add reference to CPBase.DLL, typically installed in c:\program files\kma\contensive\
        '
        '=====================================================================================
        ' 
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                returnHtml = "Visual Studio Contensive Addon - OK response"
            Catch ex As Exception
                errorReport(CP, ex, "execute")
            End Try
            Return returnHtml
        End Function
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterPageClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Private NewsletterID As Long

        Private isManager As Boolean

        Private EncodeCopyNeeded As Boolean
        'Private Main As ccWeb3.MainClass
        Private Main As Object
        Private Csv As Object

        Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
            On Error GoTo ErrorTrap

            Csv = CsvObject

            Call Init(MainObject)

            Execute = GetContent(OptionString)

            Exit Function
ErrorTrap:
            Call HandleError("PageClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
        End Function

        Public Sub Init(MainObject As Object)
            '
            Main = MainObject
            '
            Dim Common As New CommonClass
            Dim CurrentLink As String
            '
            Call Common.UpgradeAddOn(Main)
            EncodeCopyNeeded = (Main.ContentServerVersion < "3.3.947")
            '
            ' 1/1/09 - JK - always pointed to the root path
            CurrentLink = Main.ServerProtocol & Main.ServerHost & Main.ServerPath & Main.ServerPage & "?" & Main.RefreshQueryString
            'CurrentLink = Main.ServerProtocol & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & "?" & Main.RefreshQueryString
            '
            ReferLink = RequestNameRefer & "=" & Main.EncodeRequestVariable(ModifyQueryString(CurrentLink, RequestNameRefer, ""))

            isManager = Main.IsContentManager("Newsletters")

            Exit Sub
            '
ErrorTrap:
            Call HandleError("NewsLetter", "Init", Err.Number, Err.Source, Err.Description, True, False)
        End Sub
        '
        Public Function GetContent(OptionString As String) As String
            On Error GoTo ErrorTrap
            '
            Dim EditLink As String
            Dim Controls As String
            Dim UnpublishedIssueList As String
            Dim BuildDefault As Boolean
            Dim Stream As String
            Dim IssueID As Long
            Dim IssuePageID As Long
            Dim Common As New CommonClass
            Dim CS As Long
            Dim Body As BodyClass
            Dim TemplateID As Long
            Dim NewTemplateID As Long
            Dim FormID As Long
            Dim EmailID As Long
            Dim TemplateCopy As String
            Dim Copy As String
            Dim NewsletterName As String
            Dim AOPointer As Long
            Dim StyleString As String
            Dim QS As String
            '
            If Not (Main Is Nothing) Then
                NewsletterName = Main.GetAddonOption("Newsletter", OptionString)
                If NewsletterName = "" Then
                    NewsletterName = DefaultRecord
                End If
                NewsletterID = Common.GetNewsletterID(Main, NewsletterName)
                Call Main.TestPoint("PC NewsletterID After Option: " & NewsletterID)
                '
                BuildDefault = Main.GetStreamBoolean("BuildDefault")
                FormID = Main.GetStreamInteger(RequestNameFormID)
                IssuePageID = Main.GetStreamInteger(RequestNameIssuePageID)
                If IssuePageID = 0 Then
                    '
                    ' No page given, use the QS for the Issue, or get current
                    '
                    Call Main.TestPoint("GetIssueID call 4, NewsletterID=" & NewsletterID)
                    IssueID = Common.GetIssueID(Main, NewsletterID)
                Else
                    '
                    ' PageID given, get Issue from PageID (and check against Newsletter)
                    '
                    CS = Main.OpenCSContent(ContentNameNewsletterIssuePages, "(id=" & IssuePageID & ")", , , , , "NewsletterID")
                    If Main.IsCSOK(CS) Then
                        IssueID = Main.GetCSInteger(CS, "NewsletterID")
                    End If
                    Call Main.CloseCS(CS)
                    CS = Main.OpenCSContent(ContentNameNewsletterIssues, "(id=" & IssueID & ")and(Newsletterid=" & NewsletterID & ")", , , , , "ID")
                    If Not Main.IsCSOK(CS) Then
                        '
                        ' Bad Issue, reset to current issue of current newsletter
                        '
                        Call Main.TestPoint("GetIssueID call 5, NewsletterID=" & NewsletterID)
                        IssueID = Common.GetIssueID(Main, NewsletterID)
                        IssuePageID = 0
                        FormID = FormIssue
                    End If
                    Call Main.CloseCS(CS)
                End If
                Call Main.SetVisitProperty(VisitPropertyNewsletter, NewsletterID & "." & IssueID & "." & IssuePageID & "." & FormID)
                '
                Call Main.TestPoint("PageClass NLID: " & NewsletterID)
                '
                Call Common.SortCategoriesByIssue(Main, IssueID)
                '
                If FormID = FormEmail Then
                    '
                    ' Create Newsletter Email
                    '
                    If Not isManager Then
                        '
                        ' Not administrators
                        '
                        Call Main.AddUserError("Only administrators can use the Create Email feature.")
                        FormID = FormIssue
                        'ElseIf Main.PageContent = "" Then
                        '    '
                        '    ' Public Site Only (need a destination in all the links)
                        '    '
                        '    Call Main.AddUserError("This feature can only be used after the Newsletter Add-on has been placed on a page of the public website.")
                        '    FormID = FormIssue
                    Else
                        '
                        ' create email version -- use Print Version to block edit links
                        '
                        Main.ServerPagePrintVersion = True
                        EmailID = CreateEmailGetID(IssueID, NewsletterName, NewsletterID)
                        Call Main.Redirect(Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameGroupEmail) & "&id=" & EmailID & "&af=4")
                        Exit Function
                    End If
                End If
                '
                ' Create the Newsletter
                '
                If (IssueID = 0) Then
                    '
                    ' There are no current issues, diplay a message and tell the admin what to do next
                    '
                    GetContent = "<p>There are currently no published issues of this newsletter</p>"
                Else
                    If NewsletterID <> 0 Then
                        CS = Main.OpenCSContentRecord("Newsletters", NewsletterID, , , "StylesFilename,TemplateID")
                        If Main.IsCSOK(CS) Then
                            TemplateID = Main.GetCSInteger(CS, "TemplateID")
                            Call Main.AddStylesheetLink(Main.ServerProtocol & Main.ServerHost & Main.serverFilePath & Main.GetCSText(CS, "StylesFileName"))
                        End If
                        Call Main.CloseCS(CS)
                        '
                        If TemplateID <> 0 Then
                            CS = Main.OpenCSContentRecord("Newsletter Templates", TemplateID, , , "Template")
                            If Not Main.IsCSOK(CS) Then
                                '
                                ' template set, but the ID is bad
                                '
                                TemplateID = 0
                            Else
                                TemplateCopy = Main.GetCSText(CS, "Template")
                                If TemplateCopy = "" Then
                                    '
                                    ' template set, but the copy is empty
                                    '
                                    TemplateID = 0
                                End If
                            End If
                            Call Main.CloseCS(CS)
                        End If
                        '
                        If TemplateID = 0 Then
                            TemplateID = Common.GetDefaultTemplateID(Main)
                            If TemplateID <> 0 Then
                                CS = Main.OpenCSContentRecord("Newsletter Issues", IssueID)
                                If Main.IsCSOK(CS) Then
                                    Call Main.SetCS(CS, "TemplateID", TemplateID)
                                End If
                                Call Main.CloseCS(CS)
                            End If
                        End If
                        '
                        If TemplateID > 0 Then
                            CS = Main.OpenCSContentRecord("Newsletter Templates", TemplateID)
                            If Main.IsCSOK(CS) Then
                                EditLink = Main.GetCSRecordEditLink(CS)
                                'If EditLink <> "" Then
                                '    EditLink = EditLink & "(Edit this Newsletter Template)" & GetContent
                                'End If
                                GetContent = Main.GetCSText(CS, "Template")
                                'GetContent = "" _
                                '    & EditLink _
                                '    & Main.GetCSText(CS, "Template")
                                GetContent = Main.GetEditWrapper("Newsletter Template [" & Main.GetCSText(CS, "Name") & "] " & EditLink, GetContent)
                            End If
                            Call Main.CloseCS(CS)
                        End If
                    End If
                    '
                    'If IssueID <> 0 And Main.IsAdmin() Then
                    '    Copy = Main.GetRecordEditLink("Newsletter Issues", IssueID)
                    '    If Copy <> "" Then
                    '        GetContent = Copy & "(Edit this Newsletter Issue)" & GetContent
                    '    End If
                    'End If
                    '
                    If GetContent <> "" Then
                        '
                        ' There is a template, encoding it captures the BodyClass
                        '
                        If EncodeCopyNeeded Then
                            '
                            ' I must encode this because old version of contensive did not auto encode addon output
                            '
                            GetContent = Main.EncodeContent(GetContent, Main.memberID, -1, False, False, True, True, False, True)
                        Else
                            GetContent = Main.EncodeContent3(GetContent, Main.memberID, "Newsletter Templates", TemplateID, 0, False, False, True, True, False, True, "", Main.ServerProtocol & Main.ServerHost, False)
                            'GetContent = GetContent
                        End If
                    Else
                        '
                        ' No valid template, call just the body so get Archive Lists
                        '
                        Body = New BodyClass
                        Call Body.Init(Main)
                        GetContent = Body.GetContent("")
                    End If
                End If
                '
                ' List Unpublished issues for admins
                '
                If Main.IsAuthoring(ContentNameNewsletters) Then
                    '
                    ' Controls
                    '
                    Controls = ""
                    QS = Main.RefreshQueryString
                    If QS <> "" Then
                        QS = QS & "&"
                    Else
                        QS = QS & "?"
                    End If
                    If IssueID <> 0 Then
                        '
                        ' For this issue
                        '
                        Controls = Controls & "<h3>For this Issue</h3><ul>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssuePages) & "&af=4&aa=2&ad=1&wc=" & Main.EncodeRequestVariable("NewsletterID=" & IssueID) & "&" & ReferLink & """>Add a new story</a></div></li>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div></li>"
                        If (InStr(1, Main.ServerPathPage, "/admin", vbTextCompare) <> 0) Or (LCase(Main.SiteProperty_AdminURL) = LCase(Main.ServerPathPage)) Then
                            Controls = Controls & "<li><div class=""AdminLink"">Create&nbsp;email&nbsp;version (not available from admin site)</div></li>"
                        Else
                            Controls = Controls & "<li><div class=""AdminLink""><a href=""?" & QS & RequestNameFormID & "=" & FormEmail & "&" & RequestNameIssueID & "=" & IssueID & """>Create&nbsp;email&nbsp;version</a></div></li>"
                        End If
                        Controls = Controls & "</ul>"
                    End If
                    If NewsletterID <> 0 Then
                        '
                        ' For this newsletter
                        '
                        Controls = Controls & "<h3>For this Newsletter</h3><ul>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div></li>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletters) & "&id=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Edit the styles for this newsletter</a></div></li>"
                        Controls = Controls & "</ul>"
                        '
                        ' Search for unpublished versions
                        '
                        UnpublishedIssueList = Common.GetUnpublishedIssueList(Main, NewsletterID)
                        If UnpublishedIssueList <> "" Then
                            Controls = Controls & "<h3>Unpublished issues for this Newsletter</h3>"
                            Controls = Controls & UnpublishedIssueList
                        End If
                    End If
                    '
                    ' General Controls
                    '
                    Controls = Controls & "<h3>General Controls</h3><ul>"
                    Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameIssueCategories) & "&" & ReferLink & """>Edit categories</a></div></li>"
                    Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletters) & "&af=4&" & "&" & ReferLink & """>Add a new newsletter</a></div></li>"
                    Controls = Controls & "</ul>"
                    '
                    ' instructions
                    '
                    Controls = Controls _
                         & "<P>This addon can control one or many different newsletters on your site. For instance you may have a newsletter about site news and another about industry news. Each newsletter can have many issues. For instance, Site News may have a new issue every quarter, Industry News may have a new issue every month. Each issue can have many stories. The newsletter creates one page for the front cover with a list of stories, and one page per story. It also includes a navigation panel for all pages.</P>" _
                         & "<P>The layout of the newsletter is controlled with a Newsletter Template. Use HTML and the addons 'Newsletter-body only' and Newsletter-nav only' to design your look and feel.</P>" _
                         & "<P>If you will be creating an email from this newsletter, be sure to include your styles in either the newsletter template or the newsletter record.</P>" _
                         & "<P>When you view the newsletter addon for the first time, it will automatically create a 'Default' newsletter for you.</P>" _
                         & "<P>To create a new issue for this newsletter, click the 'Add a new Issue' link. The new issue will automatically appear to the publish on the publish date you set. Before the publish date only administrators can access the new issue as they add or modify stories.</P>" _
                         & "<P>To create a new newsletter, click the 'Add a new Newsletter' link. To make your new newsletter appear here, turn on Advanced Edit and click the Options icon at the top of add-on (wrench icon). Select the newsletter you want to display and hit update.</P>" _
                         & ""
                    If Controls <> "" Then
                        GetContent = GetContent & Main.GetAdminHintWrapper(Controls)
                    End If
                End If
                '
                ' Add any user errors
                '
                If Main.IsUserError Then
                    GetContent = "<div style=""padding:10px"">" & Main.GetUserError() & "</div>" & GetContent
                End If
                '
                ' Add newsletter edit wrapper
                '
                If Main.IsEditing("Newsletters") Then
                    GetContent = Main.GetEditWrapper("Newsletter [" & NewsletterName & "] " & Main.GetRecordEditLink2("Newsletters", NewsletterID, False, NewsletterName, True), GetContent)
                End If
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function CreateEmailGetID(IssueID As Long, NewsletterName As String, NewsletterID As Long) As Long

            On Error GoTo ErrorTrap

            Dim EmailAddress As String
            Dim MemberName As String
            Dim CSPointer As Long
            Dim TemplatePointer As Long
            Dim EmailPointer As Long
            Dim GroupPointer As Long
            Dim SQL As String
            Dim LocalGroupID As Long
            Dim Caption As String
            Dim GroupName As String
            Dim EmailID As Long
            Dim TemplateCopy As String
            Dim Copy As String
            Dim Stream As String
            Dim Common As New CommonClass
            Dim CS As Long
            Dim Body As BodyClass
            Dim TemplateID As Long
            Dim Pos As Long
            Dim posStart As Long
            Dim posEnd As Long
            Dim NavObj As NavClass
            Dim BodyObj As BodyClass
            Dim Styles As String

            If IssueID > 0 Then
                CS = Main.OpenCSContentRecord("Newsletters", NewsletterID)
                If Main.IsCSOK(CS) Then
                    TemplateID = Main.GetCSInteger(CS, "TemplateID")
                    Styles = Main.ReadVirtualFile(Main.GetCSText(CS, "StylesFileName"))
                End If
                Call Main.CloseCS(CS)

                If TemplateID = 0 Then
                    TemplateID = Main.GetRecordID("Newsletter Templates", "Default")
                End If

                CS = Main.OpenCSContentRecord("Newsletter Templates", TemplateID)
                If Main.IsCSOK(CS) Then
                    Copy = Main.GetCSText(CS, "Template")
                End If
                Call Main.CloseCS(CS)
            End If

            If Copy <> "" Then
                '
                ' replace-in the navigation
                ' if I call EncodeContent, it will also encode all the personalization
                '
                Pos = InStr(1, Copy, "Newsletter-nav only", vbTextCompare)
                If Pos <> 0 Then
                    posStart = InStrRev(Copy, "<AC ", Pos, vbTextCompare)
                    If posStart <> 0 Then
                        posEnd = InStr(posStart, Copy, ">")
                        If posEnd > 0 Then
                            NavObj = New NavClass
                            Call NavObj.Init(Main)
                            Copy = Mid(Copy, 1, posStart - 1) & NavObj.GetContent("newsletter=" & NewsletterName) & Mid(Copy, posEnd + 1)
                            NavObj = Nothing
                        End If
                    End If
                End If
                '
                Pos = InStr(1, Copy, "Newsletter-body only", vbTextCompare)
                If Pos <> 0 Then
                    posStart = InStrRev(Copy, "<AC ", Pos, vbTextCompare)
                    If posStart <> 0 Then
                        posEnd = InStr(posStart, Copy, ">")
                        If posEnd > 0 Then
                            BodyObj = New BodyClass
                            Call BodyObj.Init(Main)
                            Copy = Mid(Copy, 1, posStart - 1) & BodyObj.GetContent("newsletter=" & NewsletterName) & Mid(Copy, posEnd + 1)
                            BodyObj = Nothing
                        End If
                    End If
                End If

                '
                '   JF 6/23/09 - this will catch any add-ons droppped anywhere, but more importnatly in the template itself
                '
                Copy = Main.EncodeContent(Copy, Main.memberID, -1, False, False, True, True, False, True)

            Else
                '
                ' No valid template, call just the body so get Archive Lists
                '
                Body = New BodyClass
                Call Body.Init(Main)
                Copy = Body.GetContent("")
            End If

            If Styles <> "" Then
                Copy = "<style>" & Styles & "</style>" & Copy
            End If
            '
            ' Remove comments - dont know why, but emails fail with comments embedded
            '
            Dim LoopPtr As Long
            Dim StartPos As Long
            Dim EndPos As Long
            LoopPtr = 0
            Do While InStr(1, Copy, "<!--") <> 0 And LoopPtr < 100
                StartPos = InStr(1, Copy, "<!--")
                EndPos = InStr(StartPos, Copy, "-->")
                If EndPos <> 0 Then
                    Copy = Left(Copy, StartPos - 1) & Mid(Copy, EndPos + 3)
                End If
                LoopPtr = LoopPtr + 1
            Loop
            '
            ' ----- add inline styles because gmail removes style tags
            '
            Copy = Replace(Copy, "class=""Headline""", "class=""Headline"" style=""font-weight:bold;margin-top:20px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""Overview""", "class=""Overview"" style=""margin-top:10px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""PrintIcon""", "class=""PrintIcon"" style=""float:right;text-align: right; margin-top:5px;margin-bottom:10px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""EmailIcon""", "class=""EmailIcon"" style=""float:right;text-align: right; margin-top:5px;margin-bottom:10px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""ReadMore""", "class=""ReadMore"" style=""margin-top:5px;margin-bottom:10px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""NewsletterTopic""", "class=""NewsletterTopic"" style=""font-weight:bold;padding-top:15px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""NewsletterTopicStory""", "class=""NewsletterTopicStory"" style=""padding-left:20px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""GoToPageLine""", "class=""GoToPageLine"" style=""text-align:left;margin-top:10px;padding-top:10px;border-top:1px solid black;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""LinkLine""", "class=""LinkLine"" style=""padding:10px 0 0 0;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""caption""", "class=""caption"" style=""font-weight:bold;margin-top: 10px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""PageList""", "class=""PageList"" style=""margin-top: 10px;margin-left:10px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""NewsletterNavTopic""", "class=""NewsletterNavTopic"" style=""margin-top:15px;"" ", , , vbTextCompare)
            Copy = Replace(Copy, "class=""NewsletterNavTopic""", "class=""NewsletterNavTopic"" style=""margin-top:15px;"" ", , , vbTextCompare)
            '
            EmailPointer = Main.InsertCSContent(ContentNameGroupEmail)
            If Main.CSOK(EmailPointer) Then
                CreateEmailGetID = Main.GetCSInteger(EmailPointer, "ID")
                If NewsletterName = "" Then
                    NewsletterName = Main.GetRecordName(ContentNameNewsletterIssues, IssueID)
                End If
                EmailID = Main.GetCSInteger(EmailPointer, "ID")
                EmailAddress = Trim(Main.MemberEmail)
                MemberName = Main.MemberName
                If EmailAddress = "" Then
                    EmailAddress = Main.EmailAdmin
                End If
                If (EmailAddress <> "") And (MemberName <> "") Then
                    EmailAddress = """" & MemberName & """ <" & EmailAddress & ">"
                End If
                Call Main.SetCS(EmailPointer, "Name", "Newsletter " & NewsletterName)
                Call Main.SetCS(EmailPointer, "Subject", NewsletterName)
                Call Main.SetCS(EmailPointer, "FromAddress", EmailAddress)
                Call Main.SetCS(EmailPointer, "TestMemberID", Main.memberID)
                Call Main.SetCSTextFile(EmailPointer, "CopyFileName", Copy, ContentNameGroupEmail)
                Call Main.SaveCSRecord(EmailPointer)
            End If
            Call Main.CloseCS(EmailPointer)
            '
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "CreateEmailGetID", Err.Number, Err.Source, Err.Description, True, False)
        End Function

    End Class
End Namespace
