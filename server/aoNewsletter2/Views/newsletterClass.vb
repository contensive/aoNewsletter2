﻿
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.Addons.Newsletter.Controllers
Imports Contensive.BaseClasses

Namespace Views
    Public Class NewsletterClass
        ' 
        '=====================================================================================
        ''' <summary>
        ''' Newsletter Addon Interface
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Function getLegacyNewsletter(ByVal CP As CPBaseClass, NewsletterID As Integer, currentIssueID As Integer) As Object
            Dim returnHtml As String = ""
            Try
                Dim refreshQueryString As String = ""
                '
                Dim layout As New BlockClass
                Dim newsBody As String = ""
                Dim newsNav As String = ""
                '
                Dim EditLink As String
                Dim Controls As String
                Dim UnpublishedIssueList As String
                Dim BuildDefault As Boolean
                Dim IssueID As Integer
                Dim storyID As Integer
                Dim cn As New NewsletterController
                Dim cs As CPCSBaseClass = CP.CSNew()
                Dim Body As NewsletterBodyClass
                Dim nav As NewsletterNavClass
                Dim TemplateID As Integer
                Dim FormID As Integer
                Dim EmailID As Integer
                Dim TemplateCopy As String = ""
                Dim qs As String
                Dim ButtonValue As String

                Dim isManager As Boolean
                Dim ReferLink As String
                Dim currentLink As String = ""
                Dim isContentManager As Boolean = CP.User.IsContentManager("newsletters")
                Dim newsCoverItemList As String = ""
                Dim itemLayout As String = ""
                Dim itemLayoutStory As String = ""
                Dim itemLayoutCategory As String = ""
                Dim isEditing As Boolean = CP.User.IsEditingAnything()
                ' deal with this later
                Dim archiveIssueID As Integer = 0
                Dim ItemList As String = ""

                Dim footerAdBanners As String = ""
                Dim itemLayoutAdBanners As String = ""
                Dim sponsor As String = ""
                Dim publishDate As Date = Date.MinValue
                Dim tagLine As String = ""
                Dim mastheadFilename As String = ""
                Dim footerFilename As String = ""
                Dim problemList As New List(Of String)
                '
                refreshQueryString = CP.Doc.RefreshQueryString
                '
                currentLink = CP.Request.Protocol & CP.Site.DomainPrimary & CP.Request.PathPage & "?" & refreshQueryString
                ReferLink = RequestNameRefer & "=" & CP.Utils.EncodeRequestVariable(CP.Utils.ModifyLinkQueryString(currentLink, RequestNameRefer, ""))
                isManager = CP.User.IsContentManager("Newsletters")
                '
                BuildDefault = CP.Doc.GetBoolean("BuildDefault")
                FormID = CP.Doc.GetInteger(RequestNameFormID)
                storyID = CP.Doc.GetInteger(RequestNameStoryId)
                If storyID = 0 Then
                    '
                    ' No page given, use the QS for the Issue, or get current
                    '
                    Call CP.Site.TestPoint("GetIssueID call 4, NewsletterID=" & NewsletterID)
                    IssueID = NewsletterController.GetIssueID(CP, NewsletterID, currentIssueID)
                Else
                    '
                    ' PageID given, get Issue from PageID (and check against Newsletter)
                    '
                    Call cs.Open(ContentNameNewsletterStories, "(id=" & storyID & ")")
                    If cs.OK() Then
                        IssueID = cs.GetInteger("NewsletterID")
                    End If
                    Call cs.Close()
                    '
                    Call cs.Open(ContentNameNewsletterIssues, "active=1 and (id=" & IssueID & ")and(Newsletterid=" & NewsletterID & ")")
                    If Not cs.OK() Then
                        '
                        ' Bad Issue, reset to current issue of current newsletter
                        '
                        Call CP.Site.TestPoint("GetIssueID call 5, NewsletterID=" & NewsletterID)
                        IssueID = NewsletterController.GetIssueID(CP, NewsletterID, currentIssueID)
                        storyID = 0
                        FormID = FormCover
                    End If
                    Call cs.Close()
                End If
                Call CP.Visit.SetProperty(VisitPropertyNewsletter, NewsletterID & "." & IssueID & "." & storyID & "." & FormID)
                '
                Call CP.Site.TestPoint("PageClass NLID: " & NewsletterID)
                '
                Call NewsletterController.SortCategoriesByIssue(CP, IssueID)
                '
                If (isManager And (FormID = FormEmail)) Then
                    '
                    ' create email version -- use Print Version to block edit links
                    '
                    EmailID = CreateEmailGetID(CP, IssueID, NewsletterID, refreshQueryString, currentIssueID)
                    CP.Response.Redirect(CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameGroupEmail) & "&id=" & EmailID & "&af=4")
                    returnHtml = ""
                ElseIf (FormID = FormEmail) Then
                    '
                    ' Not administrators
                    '
                    Call CP.UserError.Add("Only administrators can use the Create Email feature.")
                    FormID = FormCover
                Else
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
                            Call openRecord(CP, cs, "Newsletters", NewsletterID, "StylesFilename,TemplateID,mastheadFilename,footerFilename")
                            If cs.OK() Then
                                TemplateID = cs.GetInteger("TemplateID")
                                mastheadFilename = cs.GetText("mastheadFilename")
                                footerFilename = cs.GetText("footerFilename")
                            End If
                            Call cs.Close()
                            '
                            If TemplateID <> 0 Then
                                Call openRecord(CP, cs, "Newsletter Templates", TemplateID, "Template")
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
                                TemplateID = NewsletterController.verifyDefaultTemplateGetId(CP)
                                If TemplateID <> 0 Then
                                    Call openRecord(CP, cs, "Newsletters", IssueID)
                                    If cs.OK() Then
                                        Call cs.SetField("TemplateID", TemplateID.ToString())
                                    End If
                                    Call cs.Close()
                                End If
                            End If
                            '
                            If TemplateID > 0 Then
                                Call openRecord(CP, cs, "Newsletter Templates", TemplateID)
                                If cs.OK() Then
                                    EditLink = cs.GetEditLink()
                                    TemplateCopy = cs.GetText("Template")
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
                        layout.load(TemplateCopy)
                        '
                        ' -- masthead image
                        If (Not String.IsNullOrEmpty(mastheadFilename)) Then
                            mastheadFilename = Uri.EscapeUriString(mastheadFilename)
                            layout.setClassInner("newsHeaderMasthead", "<img src=""" & CP.Http.CdnFilePathPrefix & mastheadFilename & """ style=""width:100%"" class=""banner"" />")
                        End If
                        '
                        ' -- footer image
                        If (Not String.IsNullOrEmpty(footerFilename)) Then
                            footerFilename = Uri.EscapeUriString(footerFilename)
                            layout.setClassInner("newsFooterMasthead", "<img src=""" & CP.Http.CdnFilePathPrefix & footerFilename & """ style=""width:100%"" class=""banner"" />")
                        End If
                        '
                        nav = New NewsletterNavClass
                        newsNav = layout.getClassInner("newsNav")
                        '
                        Body = New NewsletterBodyClass
                        Select Case FormID
                            Case FormSearch
                                itemLayout = layout.getClassOuter("newsSearchListItem")
                                ItemList = Body.GetSearchItemList(CP, cn, ButtonValue, IssueID, refreshQueryString, itemLayout)
                                itemLayoutAdBanners = layout.getClassOuter("adBannerItem")
                                Call layout.setClassOuter("newsSearchList", ItemList)
                                Call layout.setClassInner("newsArchive", "")
                                Call layout.setClassOuter("newsBody", "")
                                Call layout.setClassOuter("newsCover", "")
                                Call layout.setClassOuter("emailLinkToWeb", "")
                                Call layout.setClassOuter("newsIssueCaption", "")
                                Call layout.setClassInner("newsIssueSponsor", sponsor)
                                Call layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString)
                                If (String.IsNullOrEmpty(tagLine)) Then
                                    Call layout.setClassOuter("newsletterTagLine", "")
                                Else
                                    Call layout.setClassInner("newsletterTagLine", tagLine)
                                End If
                                newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID)
                            Case FormArchive
                                '
                                '
                                '
                                Dim searchForm As String = ""
                                searchForm &= "<div>"
                                searchForm &= CP.Html.InputText(RequestNameSearchKeywords)
                                searchForm &= " <input type=""submit"" id=""js-ArchiveIssuesSubmit"" name=""Button"" value="" Search ""> " 'CP.Html.Button(FormButtonViewArchives, FormButtonViewArchives)
                                searchForm &= "</div>"
                                searchForm = CP.Html.Form(searchForm, "", "", "", CP.Utils.ModifyQueryString(refreshQueryString, RequestNameFormID, FormArchive.ToString()))

                                Call layout.setClassInner("newsArchiveSearch", searchForm)
                                '
                                '
                                '
                                itemLayout = layout.getClassOuter("newsArchiveListItem")
                                ItemList = Body.GetArchiveItemList(CP, cn, ButtonValue, currentIssueID, refreshQueryString, itemLayout, NewsletterID)
                                itemLayoutAdBanners = layout.getClassOuter("adBannerItem")
                                Call layout.setClassInner("newsArchiveList", ItemList)
                                Call layout.setClassOuter("newsBody", "")
                                Call layout.setClassOuter("newsCover", "")
                                Call layout.setClassOuter("newsSearch", "")
                                Call layout.setClassOuter("emailLinkToWeb", "")
                                Call layout.setClassOuter("newsIssueCaption", "")
                                Call layout.setClassInner("newsIssueSponsor", sponsor)
                                Call layout.setClassInner("newsIssuePublishDate", "")
                                If (String.IsNullOrEmpty(tagLine)) Then
                                    Call layout.setClassOuter("newsletterTagLineRow", "")
                                Else
                                    Call layout.setClassInner("newsletterTagLine", "")
                                End If
                                newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID)
                            Case FormStory
                                newsBody = layout.getClassInner("newsBody")
                                If (String.IsNullOrEmpty(newsBody.Trim())) Then
                                    problemList.Add("The newsletter template does not contain a class with 'newsBody', required for a detail page.")
                                End If
                                newsBody = Body.GetStory(CP, cn, storyID, IssueID, refreshQueryString, newsBody, isEditing)
                                Call openRecord(CP, cs, "Newsletter Issues", IssueID)
                                If cs.OK() Then
                                    sponsor = cs.GetText("sponsor")
                                    tagLine = cs.GetText("tagLine")
                                    publishDate = cs.GetDate("publishDate")
                                End If
                                Call cs.Close()
                                itemLayoutAdBanners = layout.getClassOuter("adBannerItem")
                                Call layout.setClassInner("newsBody", newsBody)
                                Call layout.setClassOuter("newsArchive", "")
                                Call layout.setClassOuter("newsCover", "")
                                Call layout.setClassOuter("newsSearch", "")
                                Call layout.setClassOuter("emailLinkToWeb", "")
                                Call layout.setClassInner("newsIssueCaption", CP.Content.GetRecordName(ContentNameNewsletterIssues, IssueID))
                                Call layout.setClassInner("newsIssueSponsor", sponsor)
                                Call layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString)
                                If (String.IsNullOrEmpty(tagLine)) Then
                                    Call layout.setClassOuter("newsletterTagLineRow", "")
                                Else
                                    Call layout.setClassInner("newsletterTagLine", tagLine)
                                End If
                                newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID)
                            Case Else
                                '
                                ' -- Form Cover
                                FormID = FormCover
                                itemLayoutStory = layout.getClassOuter("newsCoverStoryItem")
                                itemLayoutAdBanners = layout.getClassOuter("adBannerItem")
                                itemLayoutCategory = layout.getClassOuter("newsCoverCategoryItem")
                                ItemList = Body.GetCoverContent(CP, IssueID, storyID, refreshQueryString, FormID, itemLayoutStory, itemLayoutCategory, isEditing, sponsor, publishDate, tagLine)
                                '
                                ' add footer ad banner(s)
                                '
                                If (cs.Open("newsletter Issues", "id=" & IssueID)) Then
                                    Dim adBanner As String
                                    Dim adBannerLink As String
                                    Dim bannerLayoutId As Integer
                                    Dim adBannerRowCnt As Integer = 1
                                    Dim adBannerColumnCnt As Integer = 1
                                    Dim pxColumnSpace As Integer = 0
                                    Dim pxRowSpace As Integer = 0
                                    '
                                    bannerLayoutId = cs.GetInteger("bannerLayoutId")
                                    If (bannerLayoutId > 0) Then
                                        Dim csLayout As CPCSBaseClass = CP.CSNew()
                                        If csLayout.Open("Newsletter Ad Banner Layouts", "id=" & bannerLayoutId) Then
                                            adBannerRowCnt = csLayout.GetInteger("rowCnt")
                                            adBannerColumnCnt = csLayout.GetInteger("columnCnt")
                                            pxColumnSpace = csLayout.GetInteger("pxColumnSpace")
                                            pxRowSpace = csLayout.GetInteger("pxRowSpace")
                                        End If
                                        Call csLayout.Close()
                                    End If

                                    For rowPtr As Integer = 0 To adBannerRowCnt - 1
                                        If (pxRowSpace > 0) And (rowPtr > 0) Then
                                            footerAdBanners &= "<img src=""\cclib\images\spacer.gif"" width=""10"" height=""" & pxRowSpace.ToString() & """ style=""height:" & pxRowSpace.ToString() & "px"">"
                                        End If
                                        footerAdBanners &= "<div class=""newsletterAdvertisementRow"">"
                                        For columnPtr As Integer = 0 To adBannerColumnCnt - 1
                                            If (pxColumnSpace > 0) And (columnPtr > 0) Then
                                                footerAdBanners &= "<img src=""\cclib\images\spacer.gif"" width=""" & pxColumnSpace.ToString() & """ height=""10"" style=""width:" & pxColumnSpace.ToString() & "px"">"
                                            End If
                                            Dim adPtr As Integer = (rowPtr * adBannerColumnCnt) + columnPtr
                                            adBanner = cs.GetText("adBanner" & adPtr)
                                            If (Not String.IsNullOrEmpty(adBanner)) Then
                                                adBannerLink = cs.GetText("adBannerLink" & adPtr)
                                                If (String.IsNullOrEmpty(adBannerLink)) Then
                                                    adBanner = Uri.EscapeUriString(adBanner)
                                                    footerAdBanners &= "<img src=""" & CP.Http.CdnFilePathPrefix & adBanner & """>"
                                                Else
                                                    If (adBannerLink.IndexOf("://") < 0) Then
                                                        adBannerLink = "http://" & adBannerLink
                                                    End If
                                                    adBanner = Uri.EscapeUriString(adBanner)
                                                    adBannerLink = Uri.EscapeUriString(adBannerLink)
                                                    footerAdBanners &= "<a href=""" & adBannerLink & """ target=""_blank""><img src=""" & CP.Http.CdnFilePathPrefix & adBanner & """></a>"
                                                End If
                                            End If
                                        Next
                                        footerAdBanners &= "</div>"
                                    Next
                                    ''
                                    'adBanner = cs.GetText("adBanner2")
                                    'If (Not String.IsNullOrEmpty(adBanner)) Then
                                    '    footerAdBanners &= adBanner
                                    'End If
                                    ''
                                    'adBanner = cs.GetText("adBanner3")
                                    'If (Not String.IsNullOrEmpty(adBanner)) Then
                                    '    footerAdBanners &= adBanner
                                    'End If
                                    ''
                                    'adBanner = cs.GetText("adBanner4")
                                    'If (Not String.IsNullOrEmpty(adBanner)) Then
                                    '    footerAdBanners &= adBanner
                                    'End If
                                    ''
                                    'adBanner = cs.GetText("adBanner5")
                                    'If (Not String.IsNullOrEmpty(adBanner)) Then
                                    '    footerAdBanners &= adBanner
                                    'End If
                                    ''
                                    'adBanner = cs.GetText("adBanner6")
                                    'If (Not String.IsNullOrEmpty(adBanner)) Then
                                    '    footerAdBanners &= adBanner
                                    'End If
                                    '
                                End If
                                Call cs.Close()
                                If (Not String.IsNullOrEmpty(footerAdBanners)) Then
                                    Dim adBannerLayout2 As New BlockClass()
                                    adBannerLayout2.load(footerAdBanners)
                                    adBannerLayout2.setClassInner("newsletterAdvertisements", footerAdBanners)
                                    ItemList &= adBannerLayout2.getHtml()
                                End If
                                Call layout.setClassInner("newsCoverList", ItemList)
                                Call layout.setClassOuter("newsArchive", "")
                                Call layout.setClassOuter("newsBody", "")
                                Call layout.setClassOuter("newsSearch", "")
                                Call layout.setClassOuter("emailLinkToWeb", "")
                                Call layout.setClassInner("newsIssueCaption", CP.Content.GetRecordName(ContentNameNewsletterIssues, IssueID))
                                Call layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString)
                                If (String.IsNullOrWhiteSpace(sponsor)) Then
                                    Call layout.setClassOuter("newsIssueSponsor", "")
                                Else
                                    Call layout.setClassInner("newsIssueSponsor", sponsor)
                                End If
                                If (String.IsNullOrEmpty(tagLine)) Then
                                    Call layout.setClassOuter("newsletterTagLineRow", "")
                                Else
                                    Call layout.setClassInner("newsletterTagLine", tagLine)
                                End If
                                newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID)
                        End Select
                        Call layout.setClassInner("newsNav", newsNav)
                        '
                        ' Add archive link
                        '
                        Dim newsArchiveLink As String = layout.getClassInner("newsArchiveLink")
                        newsArchiveLink = newsArchiveLink.Replace("#", CP.Utils.ModifyLinkQueryString(currentLink, "formId", FormArchive.ToString))
                        layout.setClassInner("newsArchiveLink", newsArchiveLink)
                        '
                        returnHtml = layout.getHtml()
                    End If
                    '
                    ' List Unpublished issues for admins
                    '
                    If isEditing Then
                        '
                        ' -- wrap in issue edit
                        returnHtml = CP.Content.GetEditWrapper(CP.Content.GetEditLink("newsletter issues", currentIssueID) & returnHtml)
                        '
                        ' Controls
                        '
                        Controls = ""
                        qs = refreshQueryString
                        If qs <> "" Then
                            qs = qs & "&"
                        Else
                            qs = qs & "?"
                        End If
                        If (problemList.Count > 0) Then
                            Dim controlItems As String = ""
                            For Each problem As String In problemList
                                controlItems += CP.Html.li(problem)
                            Next
                            Controls = Controls & "<h3>Problems Found on this Page</h3>"
                            Controls += CP.Html.ul(controlItems)
                        End If
                        If IssueID <> 0 Then
                            '
                            ' For this issue
                            '
                            Controls = Controls & "<h3>For this Issue</h3><ul>"
                            Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletterStories) & "&af=4&aa=2&ad=1&wc=" & CP.Utils.EncodeRequestVariable("NewsletterID=" & IssueID) & "&" & ReferLink & """>Add a new story</a></div></li>"
                            Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div></li>"
                            If (InStr(1, CP.Request.PathPage, "/admin", vbTextCompare) <> 0) Or (LCase(CP.Site.GetText("adminUrl")) = LCase(CP.Request.PathPage)) Then
                                Controls = Controls & "<li><div class=""AdminLink"">Create&nbsp;email&nbsp;version (not available from admin site)</div></li>"
                            Else
                                qs = CP.Doc.RefreshQueryString
                                qs = CP.Utils.ModifyQueryString(qs, RequestNameFormID, FormEmail.ToString())
                                qs = CP.Utils.ModifyQueryString(qs, RequestNameIssueID, IssueID.ToString())
                                Controls = Controls & "<li><div class=""AdminLink""><a href=""?" & qs & """>Create&nbsp;email&nbsp;version</a></div></li>"
                            End If
                            Controls = Controls & "</ul>"
                        End If
                        If NewsletterID <> 0 Then
                            '
                            ' For this newsletter
                            '
                            Controls = Controls & "<h3>For this Newsletter</h3><ul>"
                            Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div></li>"
                            Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletters) & "&id=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Edit this newsletter</a></div></li>"
                            Controls = Controls & "</ul>"
                            '
                            ' Search for unpublished versions
                            '
                            UnpublishedIssueList = NewsletterController.GetUnpublishedIssueList(CP, NewsletterID, cn)
                            If UnpublishedIssueList <> "" Then
                                Controls = Controls & "<h3>Unpublished issues for this Newsletter</h3>"
                                Controls = Controls & UnpublishedIssueList
                            End If
                        End If
                        '
                        ' General Controls
                        '
                        Controls = Controls & "<h3>General</h3><ul>"
                        Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameIssueCategories) & "&" & ReferLink & """>Edit categories</a></div></li>"
                        ' Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletters) & "&af=4&" & "&" & ReferLink & """>Add a new newsletter</a></div></li>"
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
                            returnHtml = returnHtml & NewsletterController.GetAdminHintWrapper(CP, Controls)
                        End If

                    End If
                    '
                    ' Add any user errors
                    '
                    If Not CP.UserError.OK Then
                        returnHtml = "<div style=""padding:10px"">" & CP.UserError.GetList() & "</div>" & returnHtml
                    End If
                    'returnHtml = GetContent(CP, refreshQueryString)
                End If
            Catch ex As Exception
                HandleError(CP, ex, "execute")
            End Try
            Call CP.Addon.ExecuteAsProcessByUniqueName("RSS Feed Process")
            Return returnHtml
        End Function
        '
        '=====================================================================================
        ' common report for this class
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
        Private Function CreateEmailGetID(ByVal cp As CPBaseClass, ByVal IssueID As Integer, ByVal NewsletterID As Integer, ByVal refreshQueryString As String, ByVal currentIssueId As Integer) As Integer
            Dim returnId As Integer = 0
            Try
                Dim NewsletterName As String
                Dim EmailAddress As String
                Dim MemberName As String
                Dim CSPointer As CPCSBaseClass = cp.CSNew()
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim templateCopy As String = ""
                Dim cn As New NewsletterController
                Dim Body As NewsletterBodyClass
                Dim webTemplateID As Integer
                Dim Nav As NewsletterNavClass
                Dim Styles As String
                Dim layout As New BlockClass
                Dim itemList As String = ""
                Dim newsNav As String = ""
                Dim emailBody As String = ""
                Dim LoopPtr As Integer
                Dim StartPos As Integer
                Dim EndPos As Integer
                Dim newsCoverStoryItem As String = ""
                Dim itemLayoutAdBanners As String = ""
                Dim newsCoverCategoryItem As String = ""
                Dim emailTemplateID As Integer = 0
                Dim updateNewsletterTemplateId As Boolean = False
                Dim templateId As Integer = 0
                Dim adBannerLink As String
                Dim mastheadFilename As String = ""
                Dim footerFilename As String = ""
                '
                If IssueID > 0 Then
                    Call openRecord(cp, cs, "Newsletters", NewsletterID)
                    If cs.OK() Then
                        webTemplateID = cs.GetInteger("TemplateID")
                        emailTemplateID = cs.GetInteger("emailTemplateID")
                        Styles = cp.CdnFiles.Read(cs.GetText("StylesFileName"))
                        mastheadFilename = cs.GetText("mastheadFilename")
                        footerFilename = cs.GetText("footerFilename")
                    End If
                    Call cs.Close()
                    '
                    templateId = emailTemplateID
                    If templateId <> 0 Then
                        '
                        ' verify it
                        Call openRecord(cp, cs, "newsletter templates", templateId)
                        If cs.OK Then
                            templateCopy = cs.GetText("Template")
                        End If
                        Call cs.Close()
                    End If
                    '
                    If String.IsNullOrEmpty(templateCopy) Then
                        '
                        ' -- no email template available, rebuild from installation file
                        templateId = NewsletterController.verifyDefaultEmailTemplateGetId(cp)
                        Call cp.Db.ExecuteNonQuery("update newsletters set emailTemplateID=" & templateId & " where id=" & NewsletterID)
                        Call openRecord(cp, cs, "newsletter templates", templateId)
                        If cs.OK Then
                            templateCopy = cs.GetText("Template")
                        End If
                        Call cs.Close()
                    End If
                    '
                    If String.IsNullOrEmpty(templateCopy) Then
                        '
                        ' -- if all else fails, use web template
                        Call openRecord(cp, cs, "Newsletter Templates", webTemplateID)
                        If cs.OK() Then
                            templateCopy = cs.GetText("Template")
                        End If
                        Call cs.Close()
                    End If
                End If
                '
                ' There is a template, encoding it captures the newsletterBodyClass
                '
                Dim sponsor As String = ""
                Dim publishDate As Date = Date.MinValue
                Dim tagLine As String = ""
                Dim emailLinkToWebHtml As String
                Dim qs As String
                '
                layout.load(templateCopy)
                If (Not String.IsNullOrEmpty(mastheadFilename)) Then
                    mastheadFilename = Uri.EscapeUriString(mastheadFilename)
                    layout.setClassInner("newsHeaderMasthead", "<img width=""100%"" src=""" & cp.Http.CdnFilePathPrefix & mastheadFilename & """ class=""banner"" />")
                End If
                If (Not String.IsNullOrEmpty(footerFilename)) Then
                    footerFilename = Uri.EscapeUriString(footerFilename)
                    layout.setClassInner("newsFooterMasthead", "<img width=""100%"" src=""" & cp.Http.CdnFilePathPrefix & footerFilename & """ class=""footer"" />")
                End If
                '
                ' set the link back to the web version
                '
                emailLinkToWebHtml = layout.getClassInner("emailLinkToWeb")
                If Not String.IsNullOrEmpty(emailLinkToWebHtml) Then
                    qs = cp.Doc.RefreshQueryString()
                    qs = cp.Utils.ModifyQueryString(qs, "issueId", IssueID.ToString())
                    emailLinkToWebHtml = emailLinkToWebHtml.Replace("href=""#""", "href=""?" & qs & """")
                    layout.setClassInner("emailLinkToWeb", emailLinkToWebHtml)
                End If
                '
                newsCoverStoryItem = layout.getClassOuter("newsCoverStoryItem")
                itemLayoutAdBanners = layout.getClassOuter("adBannerItem")
                newsCoverCategoryItem = layout.getClassOuter("newsCoverCategoryItem")
                Body = New NewsletterBodyClass
                itemList = Body.GetCoverContent(cp, IssueID, 0, refreshQueryString, FormCover, newsCoverStoryItem, newsCoverCategoryItem, False, sponsor, publishDate, tagLine)
                '
                'Call cp.Utils.AppendLogFile("createEmailGetId, 300")
                '
                ''
                '' add footer ad banner(s)
                ''
                Dim footerAdBanners As String = ""
                '
                If (cs.Open("newsletter Issues", "id=" & IssueID)) Then
                    Dim adBanner As String
                    'Dim adBannerLink As String
                    Dim bannerLayoutId As Integer
                    Dim adBannerRowCnt As Integer = 1
                    Dim adBannerColumnCnt As Integer = 1
                    Dim pxColumnSpace As Integer = 0
                    Dim pxRowSpace As Integer = 0
                    '
                    bannerLayoutId = cs.GetInteger("bannerLayoutId")
                    If (bannerLayoutId > 0) Then
                        Dim csLayout As CPCSBaseClass = cp.CSNew()
                        If csLayout.Open("Newsletter Ad Banner Layouts", "id=" & bannerLayoutId) Then
                            adBannerRowCnt = csLayout.GetInteger("rowCnt")
                            adBannerColumnCnt = csLayout.GetInteger("columnCnt")
                            pxColumnSpace = csLayout.GetInteger("pxColumnSpace")
                            pxRowSpace = csLayout.GetInteger("pxRowSpace")
                        End If
                        Call csLayout.Close()
                    End If

                    For rowPtr As Integer = 0 To adBannerRowCnt - 1
                        If (pxRowSpace > 0) And (rowPtr > 0) Then
                            footerAdBanners &= "<img src=""\cclib\images\spacer.gif"" width=""10"" height=""" & pxRowSpace.ToString() & """ style=""height:" & pxRowSpace.ToString() & "px"">"
                        End If
                        footerAdBanners &= "<div class=""newsletterAdvertisementRow"">"
                        For columnPtr As Integer = 0 To adBannerColumnCnt - 1
                            If (pxColumnSpace > 0) And (columnPtr > 0) Then
                                footerAdBanners &= "<img src=""\cclib\images\spacer.gif"" width=""" & pxColumnSpace.ToString() & """ height=""10"" style=""width:" & pxColumnSpace.ToString() & "px"">"
                            End If
                            Dim adPtr As Integer = (rowPtr * adBannerColumnCnt) + columnPtr
                            adBanner = cs.GetText("adBanner" & adPtr)
                            If (Not String.IsNullOrEmpty(adBanner)) Then
                                adBannerLink = cs.GetText("adBannerLink" & adPtr)
                                If (String.IsNullOrEmpty(adBannerLink)) Then
                                    adBanner = Uri.EscapeUriString(adBanner)
                                    footerAdBanners &= "<img src=""" & cp.Http.CdnFilePathPrefix & adBanner & """>"
                                Else
                                    If (adBannerLink.IndexOf("://") < 0) Then
                                        adBannerLink = "http://" & adBannerLink
                                    End If
                                    adBanner = Uri.EscapeUriString(adBanner)
                                    adBannerLink = Uri.EscapeUriString(adBannerLink)
                                    footerAdBanners &= "<a href=""" & adBannerLink & """ target=""_blank""><img src=""" & cp.Http.CdnFilePathPrefix & adBanner & """></a>"
                                End If
                            End If
                        Next
                        footerAdBanners &= "</div>"
                    Next
                End If
                Call cs.Close()
                If (Not String.IsNullOrEmpty(footerAdBanners)) Then
                    Dim adBannerLayout As New BlockClass()
                    adBannerLayout.load(itemLayoutAdBanners)
                    adBannerLayout.setClassInner("newsletterAdvertisements", footerAdBanners)
                    itemList &= adBannerLayout.getHtml()
                End If
                '
                newsNav = layout.getClassInner("newsNav")
                Nav = New NewsletterNavClass
                newsNav = Nav.GetNav(cp, IssueID, NewsletterID, False, 0, newsNav, currentIssueId)
                '
                Call layout.setClassInner("newsNav", newsNav)
                Call layout.setClassInner("newsCoverList", itemList)
                Call layout.setClassOuter("newsBody", "")
                Call layout.setClassOuter("newsArchive", "")
                Call layout.setClassOuter("newsSearch", "")
                Call layout.setClassInner("newsIssueCaption", cp.Content.GetRecordName(ContentNameNewsletterIssues, IssueID))
                Call layout.setClassInner("newsIssueSponsor", sponsor)
                Call layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString)
                If (String.IsNullOrEmpty(tagLine)) Then
                    '
                    Call layout.setClassOuter("newsletterTagLineRow", "")
                Else
                    '
                    Call layout.setClassInner("newsletterTagLine", tagLine)
                End If
                '
                ' Add archive link
                '
                Dim newsArchiveLink As String = layout.getClassInner("newsArchiveLink")
                newsArchiveLink = newsArchiveLink.Replace("#", cp.Utils.ModifyLinkQueryString("?" & refreshQueryString, "formId", FormArchive.ToString))
                layout.setClassInner("newsArchiveLink", newsArchiveLink)
                '
                emailBody = layout.getHtml()
                '
                ' Remove comments - dont know why, but emails fail with comments embedded
                '
                LoopPtr = 0
                Do While InStr(1, emailBody, "<!--") <> 0 And LoopPtr < 100
                    StartPos = InStr(1, emailBody, "<!--")
                    EndPos = InStr(StartPos, emailBody, "-->")
                    If EndPos <> 0 Then
                        emailBody = Left(emailBody, StartPos - 1) & Mid(emailBody, EndPos + 3)
                    End If
                    LoopPtr = LoopPtr + 1
                Loop
                '
                Call cs.Insert(ContentNameGroupEmail)
                If cs.OK Then
                    returnId = cs.GetInteger("ID")
                    NewsletterName = cp.Content.GetRecordName(ContentNameNewsletterIssues, IssueID)
                    EmailAddress = Trim(cp.User.Email)
                    MemberName = cp.User.Name
                    If (EmailAddress <> "") And (MemberName <> "") Then
                        EmailAddress = """" & MemberName & """ <" & EmailAddress & ">"
                    End If
                    Call cs.SetField("Name", "Newsletter " & NewsletterName)
                    Call cs.SetField("Subject", NewsletterName)
                    Call cs.SetField("FromAddress", EmailAddress)
                    Call cs.SetField("TestMemberID", cp.User.Id.ToString())
                    Call cs.SetField("CopyFileName", emailBody)
                    Call cs.Save()
                End If
                Call cs.Close()
                '
                'Call cp.Utils.AppendLogFile("createEmailGetId, 999")
                '
            Catch ex As Exception
                Call HandleError(cp, ex, "CreateEmailGetID")
            End Try
            Return returnId
        End Function

    End Class
End Namespace
