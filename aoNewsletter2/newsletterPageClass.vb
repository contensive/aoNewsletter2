﻿
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterPageClass
        '
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub HandleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterPageClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Public Function GetContent(cp As CPBaseClass) As String
            Try
                '
                Dim EditLink As String
                Dim Controls As String
                Dim UnpublishedIssueList As String
                Dim BuildDefault As Boolean
                Dim Stream As String
                Dim IssueID As Integer
                Dim IssuePageID As Integer
                Dim Common As New newsletterCommonClass
                Dim cs As cpcsBaseClass = cp.csNew()
                Dim Body As newsletterBodyClass
                Dim TemplateID As Integer
                Dim NewTemplateID As Integer
                Dim FormID As Integer
                Dim EmailID As Integer
                Dim TemplateCopy As String
                Dim Copy As String
                Dim NewsletterName As String
                Dim AOPointer As Integer
                Dim StyleString As String
                Dim QS As String
                '
                Dim NewsletterID As Integer
                Dim isManager As Boolean
                Dim ReferLink As String
                Dim currentLink As String = ""
                '
                currentLink = cp.Request.Protocol & cp.Site.DomainPrimary & cp.Request.PathPage & "?" & cp.Doc.RefreshQueryString
                ReferLink = RequestNameRefer & "=" & cp.Utils.EncodeRequestVariable(cp.Utils.ModifyQueryString(currentLink, RequestNameRefer, ""))
                isManager = cp.User.IsContentManager("Newsletters")
                '
                NewsletterName = cp.Doc.gettext("Newsletter")
                If NewsletterName = "" Then
                    NewsletterName = DefaultRecord
                End If
                NewsletterID = Common.GetNewsletterID(cp, NewsletterName)
                Call cp.Site.TestPoint("PC NewsletterID After Option: " & NewsletterID)
                '
                BuildDefault = cp.Doc.GetBoolean("BuildDefault")
                FormID = cp.doc.getInteger(RequestNameFormID)
                IssuePageID = cp.doc.getInteger(RequestNameIssuePageID)
                If IssuePageID = 0 Then
                    '
                    ' No page given, use the QS for the Issue, or get current
                    '
                    Call cp.Site.TestPoint("GetIssueID call 4, NewsletterID=" & NewsletterID)
                    IssueID = Common.GetIssueID(cp, NewsletterID)
                Else
                    '
                    ' PageID given, get Issue from PageID (and check against Newsletter)
                    '
                    Call cs.open(ContentNameNewsletterIssuePages, "(id=" & IssuePageID & ")", , , , , "NewsletterID")
                    If cs.ok() Then
                        IssueID = cs.getInteger("NewsletterID")
                    End If
                    Call cs.close()
                    '
                    Call cs.open(ContentNameNewsletterIssues, "(id=" & IssueID & ")and(Newsletterid=" & NewsletterID & ")", , , , , "ID")
                    If Not cs.ok() Then
                        '
                        ' Bad Issue, reset to current issue of current newsletter
                        '
                        Call cp.Site.TestPoint("GetIssueID call 5, NewsletterID=" & NewsletterID)
                        IssueID = Common.GetIssueID(cp, NewsletterID)
                        IssuePageID = 0
                        FormID = FormIssue
                    End If
                    Call cs.close()
                End If
                Call cp.Site.SetProperty(VisitPropertyNewsletter, NewsletterID & "." & IssueID & "." & IssuePageID & "." & FormID)
                '
                Call cp.Site.TestPoint("PageClass NLID: " & NewsletterID)
                '
                Call Common.SortCategoriesByIssue(cp, IssueID)
                '
                If FormID = FormEmail Then
                    '
                    ' Create Newsletter Email
                    '
                    If Not isManager Then
                        '
                        ' Not administrators
                        '
                        Call cp.UserError.Add("Only administrators can use the Create Email feature.")
                        FormID = FormIssue
                        'ElseIf Main.PageContent = "" Then
                        '    '
                        '    ' Public Site Only (need a destination in all the links)
                        '    '
                        '    Call cp.UserError.Add("This feature can only be used after the Newsletter Add-on has been placed on a page of the public website.")
                        '    FormID = FormIssue
                    Else
                        '
                        ' create email version -- use Print Version to block edit links
                        ' ????? 
                        '
                        'Main.ServerPagePrintVersion = True
                        EmailID = CreateEmailGetID(cp, IssueID, NewsletterName, NewsletterID)
                        cp.Response.Redirect(cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameGroupEmail) & "&id=" & EmailID & "&af=4")
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
                        Call cs.OpenRecord("Newsletters", NewsletterID, "StylesFilename,TemplateID")
                        If cs.OK() Then
                            TemplateID = cs.GetInteger("TemplateID")
                            Call cp.Doc.AddHeadStyleLink(cp.Request.Protocol & cp.Site.DomainPrimary & cp.Site.FilePath & cs.getText("StylesFileName"))
                        End If
                        Call cs.Close()
                        '
                        If TemplateID <> 0 Then
                            Call cs.OpenRecord("Newsletter Templates", TemplateID, "Template")
                            If Not cs.OK() Then
                                '
                                ' template set, but the ID is bad
                                '
                                TemplateID = 0
                            Else
                                TemplateCopy = cs.getText("Template")
                                If TemplateCopy = "" Then
                                    '
                                    ' template set, but the copy is empty
                                    '
                                    TemplateID = 0
                                End If
                            End If
                            Call cs.Close()
                        End If
                        '
                        If TemplateID = 0 Then
                            TemplateID = Common.GetDefaultTemplateID(cp)
                            If TemplateID <> 0 Then
                                Call cs.OpenRecord("Newsletter Issues", IssueID)
                                If cs.OK() Then
                                    Call cs.SetField("TemplateID", TemplateID)
                                End If
                                Call cs.Close()
                            End If
                        End If
                        '
                        If TemplateID > 0 Then
                            Call cs.OpenRecord("Newsletter Templates", TemplateID)
                            If cs.OK() Then
                                EditLink = cs.GetEditLink()
                                GetContent = cs.GetText("Template")
                                GetContent = Common.GetEditWrapper(cp, "Newsletter Template [" & cs.GetText("Name") & "] " & EditLink, GetContent)
                            End If
                            Call cs.Close()
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
                        ' There is a template, encoding it captures the newsletterBodyClass
                        '

                        GetContent = cp.Utils.EncodeContentForWeb(GetContent, "Newsletter Templates", TemplateID)
                    Else
                        '
                        ' No valid template, call just the body so get Archive Lists
                        '
                        Body = New newsletterBodyClass
                        Call Body.Init(cp)
                        GetContent = Body.GetContent(cp, "")
                    End If
                End If
                '
                ' List Unpublished issues for admins
                '
                If cp.User.IsAuthoring(ContentNameNewsletters) Then
                    '
                    ' Controls
                    '
                    Controls = ""
                    QS = cp.Doc.RefreshQueryString
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
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameNewsletterIssuePages) & "&af=4&aa=2&ad=1&wc=" & cp.Utils.EncodeRequestVariable("NewsletterID=" & IssueID) & "&" & ReferLink & """>Add a new story</a></div></li>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div></li>"
                        If (InStr(1, cp.Request.PathPage, "/admin", vbTextCompare) <> 0) Or (LCase(cp.Site.GetText("adminUrl")) = LCase(cp.Request.PathPage)) Then
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
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div></li>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameNewsletters) & "&id=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Edit the styles for this newsletter</a></div></li>"
                        Controls = Controls & "</ul>"
                        '
                        ' Search for unpublished versions
                        '
                        UnpublishedIssueList = Common.GetUnpublishedIssueList(cp, NewsletterID)
                        If UnpublishedIssueList <> "" Then
                            Controls = Controls & "<h3>Unpublished issues for this Newsletter</h3>"
                            Controls = Controls & UnpublishedIssueList
                        End If
                    End If
                    '
                    ' General Controls
                    '
                    Controls = Controls & "<h3>General Controls</h3><ul>"
                    Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameIssueCategories) & "&" & ReferLink & """>Edit categories</a></div></li>"
                    Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.Site.GetText("adminUrl") & "?cid=" & cp.Content.GetID(ContentNameNewsletters) & "&af=4&" & "&" & ReferLink & """>Add a new newsletter</a></div></li>"
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
                        GetContent = GetContent & Common.GetAdminHintWrapper(cp, Controls)
                    End If
                End If
                '
                ' Add any user errors
                '
                If Not cp.UserError.OK Then
                    GetContent = "<div style=""padding:10px"">" & cp.UserError.GetList() & "</div>" & GetContent
                End If
                '
                ' Add newsletter edit wrapper
                '
                If cp.User.IsEditing("Newsletters") Then
                    GetContent = Common.GetEditWrapper(cp, "Newsletter [" & NewsletterName & "] " & cp.Content.GetEditLink("Newsletters", NewsletterID, False, NewsletterName, True), GetContent)
                End If
            Catch ex As Exception
                'Call HandleError(cp, ex, "GetContent")
            End Try
            '
            'Exit Function
        End Function
        '
        '
        '
        Private Function CreateEmailGetID(cp As CPBaseClass, IssueID As Integer, NewsletterName As String, NewsletterID As Integer) As Integer

            'On Error GoTo ErrorTrap

            Dim EmailAddress As String
            Dim MemberName As String
            Dim CSPointer As CPCSBaseClass = cp.CSNew()
            Dim TemplatePointer As Integer
            Dim EmailPointer As Integer
            Dim GroupPointer As Integer
            Dim SQL As String
            Dim LocalGroupID As Integer
            Dim Caption As String
            Dim GroupName As String
            Dim EmailID As Integer
            Dim TemplateCopy As String
            Dim Copy As String
            Dim Stream As String
            Dim Common As New newsletterCommonClass
            Dim cs As cpcsBaseClass = cp.csNew()
            Dim Body As newsletterBodyClass
            Dim TemplateID As Integer
            Dim Pos As Integer
            Dim posStart As Integer
            Dim posEnd As Integer
            Dim NavObj As newsletterNavClass
            Dim BodyObj As newsletterBodyClass
            Dim Styles As String

            If IssueID > 0 Then
                Call cs.OpenRecord("Newsletters", NewsletterID)
                If cs.ok() Then
                    TemplateID = cs.getInteger("TemplateID")
                    Styles = cp.File.ReadVirtual(cs.GetText("StylesFileName"))
                End If
                Call cs.close()

                If TemplateID = 0 Then
                    TemplateID = cp.Content.GetRecordID("Newsletter Templates", "Default")
                End If

                Call cs.OpenRecord("Newsletter Templates", TemplateID)
                If cs.ok() Then
                    Copy = cs.getText("Template")
                End If
                Call cs.close()
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
                            NavObj = New newsletterNavClass
                            Copy = Mid(Copy, 1, posStart - 1) & NavObj.GetContent(cp, "newsletter=" & NewsletterName) & Mid(Copy, posEnd + 1)
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
                            BodyObj = New newsletterBodyClass
                            Call BodyObj.Init(Main)
                            Copy = Mid(Copy, 1, posStart - 1) & BodyObj.GetContent(cp, "newsletter=" & NewsletterName) & Mid(Copy, posEnd + 1)
                            BodyObj = Nothing
                        End If
                    End If
                End If

                '
                '   JF 6/23/09 - this will catch any add-ons droppped anywhere, but more importnatly in the template itself
                '
                Copy = Main.EncodeContent(Copy, cp.user.id, -1, False, False, True, True, False, True)

            Else
                '
                ' No valid template, call just the body so get Archive Lists
                '
                Body = New newsletterBodyClass
                Call Body.Init(Main)
                Copy = Body.GetContent(cp, "")
            End If

            If Styles <> "" Then
                Copy = "<style>" & Styles & "</style>" & Copy
            End If
            '
            ' Remove comments - dont know why, but emails fail with comments embedded
            '
            Dim LoopPtr As Integer
            Dim StartPos As Integer
            Dim EndPos As Integer
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
                CreateEmailGetID = cs.getInteger(EmailPointer, "ID")
                If NewsletterName = "" Then
                    NewsletterName = Main.GetRecordName(ContentNameNewsletterIssues, IssueID)
                End If
                EmailID = cs.getInteger(EmailPointer, "ID")
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
                Call Main.SetCS(EmailPointer, "TestMemberID", cp.user.id)
                Call Main.SetCSTextFile(EmailPointer, "CopyFileName", Copy, ContentNameGroupEmail)
                Call Main.SaveCSRecord(EmailPointer)
            End If
            Call Main.CloseCS(EmailPointer)
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "CreateEmailGetID")
        End Function

    End Class
End Namespace
