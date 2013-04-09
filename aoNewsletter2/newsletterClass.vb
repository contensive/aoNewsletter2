
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterClass
        Inherits AddonBaseClass
        '
        '=====================================================================================
        ' 
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Dim refreshQueryString As String = ""
                '
                Dim layout As CPBlockBaseClass = CP.BlockNew()
                Dim newsBody As String = ""
                Dim newsNav As String = ""
                '
                Dim EditLink As String
                Dim Controls As String
                Dim UnpublishedIssueList As String
                Dim BuildDefault As Boolean
                Dim IssueID As Integer
                Dim storyID As Integer
                Dim cn As New newsletterCommonClass
                Dim cs As CPCSBaseClass = CP.CSNew()
                Dim Body As newsletterBodyClass
                Dim nav As newsletterNavClass
                Dim TemplateID As Integer
                Dim FormID As Integer
                Dim EmailID As Integer
                Dim TemplateCopy As String = ""
                Dim NewsletterName As String
                Dim QS As String
                Dim ButtonValue As String
                Dim NewsletterID As Integer
                Dim isManager As Boolean
                Dim ReferLink As String
                Dim currentLink As String = ""
                Dim isContentManager As Boolean = CP.User.IsContentManager("newsletters")
                Dim newsCoverItemList As String = ""
                Dim newsArchiveList As String = ""
                Dim newsCoverStoryItem As String = ""
                Dim newsCoverCategoryItem As String = ""
                ' deal with this later
                Dim archiveIssueID As Integer = 0
                '
                refreshQueryString = CP.Doc.RefreshQueryString
                '
                currentLink = CP.Request.Protocol & CP.Site.DomainPrimary & CP.Request.PathPage & "?" & refreshQueryString
                ReferLink = RequestNameRefer & "=" & CP.Utils.EncodeRequestVariable(CP.Utils.ModifyQueryString(currentLink, RequestNameRefer, ""))
                isManager = CP.User.IsContentManager("Newsletters")
                '
                NewsletterName = CP.Doc.GetText("Newsletter")
                If NewsletterName = "" Then
                    NewsletterName = DefaultRecord
                End If
                NewsletterID = cn.GetNewsletterID(CP, NewsletterName)
                Call CP.Site.TestPoint("PC NewsletterID After Option: " & NewsletterID)
                '
                BuildDefault = CP.Doc.GetBoolean("BuildDefault")
                FormID = CP.Doc.GetInteger(RequestNameFormID)
                storyID = CP.Doc.GetInteger(RequestNameStoryId)
                If storyID = 0 Then
                    '
                    ' No page given, use the QS for the Issue, or get current
                    '
                    Call CP.Site.TestPoint("GetIssueID call 4, NewsletterID=" & NewsletterID)
                    IssueID = cn.GetIssueID(CP, NewsletterID)
                Else
                    '
                    ' PageID given, get Issue from PageID (and check against Newsletter)
                    '
                    Call cs.Open(ContentNameNewsletterStories, "(id=" & storyID & ")", , , , , "NewsletterID")
                    If cs.OK() Then
                        IssueID = cs.GetInteger("NewsletterID")
                    End If
                    Call cs.Close()
                    '
                    Call cs.Open(ContentNameNewsletterIssues, "(id=" & IssueID & ")and(Newsletterid=" & NewsletterID & ")", , , , , "ID")
                    If Not cs.OK() Then
                        '
                        ' Bad Issue, reset to current issue of current newsletter
                        '
                        Call CP.Site.TestPoint("GetIssueID call 5, NewsletterID=" & NewsletterID)
                        IssueID = cn.GetIssueID(CP, NewsletterID)
                        storyID = 0
                        FormID = FormCover
                    End If
                    Call cs.Close()
                End If
                Call CP.Site.SetProperty(VisitPropertyNewsletter, NewsletterID & "." & IssueID & "." & storyID & "." & FormID)
                '
                Call CP.Site.TestPoint("PageClass NLID: " & NewsletterID)
                '
                Call cn.SortCategoriesByIssue(CP, IssueID)
                '
                If FormID = FormEmail Then
                    '
                    ' Create Newsletter Email
                    '
                    If Not isManager Then
                        '
                        ' Not administrators
                        '
                        Call CP.UserError.Add("Only administrators can use the Create Email feature.")
                        FormID = FormCover
                    Else
                        '
                        ' create email version -- use Print Version to block edit links
                        ' ????? 
                        '
                        'Main.ServerPagePrintVersion = True
                        EmailID = CreateEmailGetID(CP, IssueID, NewsletterName, NewsletterID, refreshQueryString)
                        CP.Response.Redirect(CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameGroupEmail) & "&id=" & EmailID & "&af=4")
                        Return ""
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
                    returnHtml = "<p>There are currently no published issues of this newsletter</p>"
                Else
                    If NewsletterID <> 0 Then
                        Call cs.OpenRecord("Newsletters", NewsletterID, "StylesFilename,TemplateID")
                        If cs.OK() Then
                            TemplateID = cs.GetInteger("TemplateID")
                            Call CP.Doc.AddHeadStyleLink(CP.Request.Protocol & CP.Site.DomainPrimary & CP.Site.FilePath & cs.GetText("StylesFileName"))
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
                                TemplateCopy = cs.GetText("Template")
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
                            TemplateID = cn.GetDefaultTemplateID(CP)
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
                                TemplateCopy = cs.GetText("Template")
                                'returnHtml = cn.GetEditWrapper(cp, "Newsletter Template [" & cs.GetText("Name") & "] " & EditLink, returnHtml)
                            End If
                            Call cs.Close()
                        End If
                    End If
                    '
                    ' Process forms
                    '
                    ButtonValue = CP.Doc.GetText("Button")
                    Select Case FormID
                        Case FormArchive
                            Select Case ButtonValue
                                Case FormButtonViewNewsLetter
                                    '
                                    ' Archive form pressing the view button
                                    '
                                    FormID = FormCover
                            End Select
                    End Select
                    '
                    ' Dispay the form
                    '
                    '
                    If TemplateCopy = "" Then
                        '
                        ' create default string 
                        '
                    End If
                    '
                    Call layout.Load(TemplateCopy)
                    '
                    nav = New newsletterNavClass
                    newsNav = layout.GetInner(".newsNav")
                    '
                    Body = New newsletterBodyClass
                    Select Case FormID
                        Case FormArchive
                            newsArchiveList = layout.GetInner(".newsArchiveList")
                            newsArchiveList = Body.GetArchiveList(CP, ButtonValue, IssueID, refreshQueryString, newsArchiveList)
                            Call layout.SetInner(".newsArchiveList", newsArchiveList)
                            Call layout.SetOuter(".newsBody", "")
                            Call layout.SetOuter(".newsCoverList", "")
                            newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav)
                        Case FormDetails
                            newsBody = layout.GetInner(".newsBody")
                            newsBody = Body.GetNewsletterBodyDetails(CP, cn, storyID, IssueID, refreshQueryString, newsBody)
                            Call layout.SetInner(".newsBody", newsBody)
                            Call layout.SetOuter(".newsArchiveList", "")
                            Call layout.SetOuter(".newsCoverList", "")
                            newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav)
                        Case Else
                            FormID = FormCover
                            newsCoverStoryItem = layout.GetOuter(".newsCoverStoryItem")
                            newsCoverCategoryItem = layout.GetOuter(".newsCoverCategoryItem")
                            newsCoverItemList = Body.GetNewsletterCover(CP, IssueID, storyID, refreshQueryString, FormID, newsCoverStoryItem, newsCoverCategoryItem)
                            Call layout.SetInner(".newsCoverList", newsCoverItemList)
                            Call layout.SetOuter(".newsArchiveList", "")
                            Call layout.SetOuter(".newsBody", "")
                            newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav)
                    End Select
                    Call layout.SetInner(".newsNav", newsNav)
                    returnHtml = layout.GetHtml()
                End If
                '
                ' List Unpublished issues for admins
                '
                If CP.User.IsAuthoring(ContentNameNewsletters) Then
                    '
                    ' Controls
                    '
                    Controls = ""
                    QS = refreshQueryString
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
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & CP.Site.DomainPrimary & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletterStories) & "&af=4&aa=2&ad=1&wc=" & CP.Utils.EncodeRequestVariable("NewsletterID=" & IssueID) & "&" & ReferLink & """>Add a new story</a></div></li>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & CP.Site.DomainPrimary & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div></li>"
                        If (InStr(1, CP.Request.PathPage, "/admin", vbTextCompare) <> 0) Or (LCase(CP.Site.GetText("adminUrl")) = LCase(CP.Request.PathPage)) Then
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
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & CP.Site.DomainPrimary & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div></li>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & CP.Site.DomainPrimary & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletters) & "&id=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Edit this newsletter</a></div></li>"
                        Controls = Controls & "</ul>"
                        '
                        ' Search for unpublished versions
                        '
                        UnpublishedIssueList = cn.GetUnpublishedIssueList(CP, NewsletterID, cn)
                        If UnpublishedIssueList <> "" Then
                            Controls = Controls & "<h3>Unpublished issues for this Newsletter</h3>"
                            Controls = Controls & UnpublishedIssueList
                        End If
                    End If
                    '
                    ' General Controls
                    '
                    Controls = Controls & "<h3>General Controls</h3><ul>"
                    Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & CP.Site.DomainPrimary & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameIssueCategories) & "&" & ReferLink & """>Edit categories</a></div></li>"
                    Controls = Controls & "<li><div class=""AdminLink""><a href = ""http://" & CP.Site.DomainPrimary & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletters) & "&af=4&" & "&" & ReferLink & """>Add a new newsletter</a></div></li>"
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
                        returnHtml = returnHtml & cn.GetAdminHintWrapper(CP, Controls)
                    End If
                End If
                '
                ' Add any user errors
                '
                If Not CP.UserError.OK Then
                    returnHtml = "<div style=""padding:10px"">" & CP.UserError.GetList() & "</div>" & returnHtml
                End If
                'returnHtml = GetContent(CP, refreshQueryString)
            Catch ex As Exception
                handleError(CP, ex, "execute")
            End Try
            Return returnHtml
        End Function
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
        '
        '
        Private Function CreateEmailGetID(ByVal cp As CPBaseClass, ByVal IssueID As Integer, ByVal NewsletterName As String, ByVal NewsletterID As Integer, ByVal refreshQueryString As String) As Integer
            Dim returnId As Integer = 0
            Try
                Dim EmailAddress As String
                Dim MemberName As String
                Dim CSPointer As CPCSBaseClass = cp.CSNew()
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim SQL As String
                Dim EmailID As Integer
                Dim templateCopy As String = ""
                Dim cn As New newsletterCommonClass
                Dim Body As newsletterBodyClass
                Dim TemplateID As Integer
                Dim Pos As Integer
                Dim posStart As Integer
                Dim posEnd As Integer
                Dim Nav As newsletterNavClass
                'Dim BodyObj As newsletterBodyClass
                Dim Styles As String
                Dim layout As CPBlockBaseClass = cp.BlockNew()
                Dim newsBody As String = ""
                Dim newsNav As String = ""
                Dim emailBody As String = ""
                Dim LoopPtr As Integer
                Dim StartPos As Integer
                Dim EndPos As Integer
                Dim newsCoverStoryItem As String = ""
                Dim newsCoverCategoryItem As String = ""
                '
                If IssueID > 0 Then
                    Call cs.OpenRecord("Newsletters", NewsletterID)
                    If cs.OK() Then
                        TemplateID = cs.GetInteger("TemplateID")
                        Styles = cp.File.ReadVirtual(cs.GetText("StylesFileName"))
                    End If
                    Call cs.Close()

                    If TemplateID = 0 Then
                        TemplateID = cp.Content.GetRecordID("Newsletter Templates", "Default")
                    End If

                    Call cs.OpenRecord("Newsletter Templates", TemplateID)
                    If cs.OK() Then
                        templateCopy = cs.GetText("Template")
                    End If
                    Call cs.Close()
                End If
                If templateCopy = "" Then
                    '
                    ' fix somehow
                    '
                End If
                '
                ' There is a template, encoding it captures the newsletterBodyClass
                '
                Call layout.Load(templateCopy)
                '
                newsBody = layout.GetInner(".newsBody")
                newsCoverStoryItem = layout.GetOuter(".newsCoverStoryList")
                newsCoverCategoryItem = layout.GetOuter(".newsCoverCategoryItem")
                Body = New newsletterBodyClass
                newsBody = Body.GetNewsletterCover(cp, IssueID, 0, refreshQueryString, FormCover, newsCoverStoryItem, newsCoverCategoryItem)
                '
                newsNav = layout.GetInner(".newsNav")
                Nav = New newsletterNavClass
                newsNav = Nav.GetNav(cp, IssueID, NewsletterID, False, 0, newsNav)
                '
                Call layout.SetInner(".newsBody", newsBody)
                Call layout.SetInner(".newsNav", newsNav)
                emailBody = layout.GetHtml()
                '
                ' Remove comments - dont know why, but emails fail with comments embedded
                '
                LoopPtr = 0
                Do While InStr(1, templateCopy, "<!--") <> 0 And LoopPtr < 100
                    StartPos = InStr(1, templateCopy, "<!--")
                    EndPos = InStr(StartPos, templateCopy, "-->")
                    If EndPos <> 0 Then
                        templateCopy = Left(templateCopy, StartPos - 1) & Mid(templateCopy, EndPos + 3)
                    End If
                    LoopPtr = LoopPtr + 1
                Loop
                '
                Call cs.Insert(ContentNameGroupEmail)
                If cs.OK Then
                    CreateEmailGetID = cs.GetInteger("ID")
                    If NewsletterName = "" Then
                        NewsletterName = cp.Content.GetRecordName(ContentNameNewsletterIssues, IssueID)
                    End If
                    EmailID = cs.GetInteger("ID")
                    EmailAddress = Trim(cp.User.Email)
                    MemberName = cp.User.Name
                    If (EmailAddress <> "") And (MemberName <> "") Then
                        EmailAddress = """" & MemberName & """ <" & EmailAddress & ">"
                    End If
                    Call cs.SetField("Name", "Newsletter " & NewsletterName)
                    Call cs.SetField("Subject", NewsletterName)
                    Call cs.SetField("FromAddress", EmailAddress)
                    Call cs.SetField("TestMemberID", cp.User.Id)
                    Call cs.SetField("CopyFileName", templateCopy)
                    Call cs.Save()
                End If
                Call cs.Close()
            Catch ex As Exception
                Call HandleError(cp, ex, "CreateEmailGetID")
            End Try
            Return returnId
        End Function

    End Class
End Namespace
