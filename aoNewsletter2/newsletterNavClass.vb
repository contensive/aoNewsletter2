
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    '
    Public Class newsletterNavClass
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub handleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterNavClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Friend Function GetNav(ByVal cp As CPBaseClass, ByVal issueid As Integer, ByVal NewsletterID As Integer, ByVal isContentManager As Boolean, ByVal FormID As Integer, ByVal newsNav As String) As String
            Dim returnHtml As String = ""
            Try
                Dim layout As CPBlockBaseClass = cp.BlockNew()
                Dim repeatItem As CPBlockBaseClass = cp.BlockNew()
                '
                Dim currentIssueId As Integer = 0
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CSPointer As CPCSBaseClass = cp.CSNew()
                Dim ThisSQL As String
                Dim Controls As String
                Dim WorkingStoryId As Integer
                Dim NavSQL As String
                Dim CategoryName As String
                Dim PreviousCategoryName As String = ""
                Dim cn As New newsletterCommonClass
                Dim AccessString As String
                Dim CategoryID As Integer
                Dim QS As String
                Dim ArticleCount As Integer
                Dim newsNavCaption As String
                Dim newsNavList As String
                Dim newsNavItem As String
                Dim storyCaption As String
                Dim repeatList As String = ""

                '
                Call layout.Load(newsNav)
                newsNavItem = layout.GetOuter(".newsNavItem")
                '
                Call repeatItem.Load(newsNavItem)
                QS = cp.Doc.RefreshQueryString
                QS = cp.Utils.ModifyQueryString(QS, RequestNameIssueID, CStr(issueid), True)
                QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormCover, True)
                Call repeatItem.SetInner(".newsNavItemCaption", "Home")
                repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                '
                NavSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder, NIC.Name AS CategoryName"
                NavSQL = NavSQL & " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR"
                NavSQL = NavSQL & " Where (NIC.ID = NIR.CategoryID)"
                NavSQL = NavSQL & " AND (NIR.NewsletterIssueID=" & issueid & ")"
                NavSQL = NavSQL & " AND (NIC.Active<>0)"
                NavSQL = NavSQL & " AND (NIR.Active<>0)"
                NavSQL = NavSQL & " ORDER BY NIR.SortOrder,NIC.name"
                '
                Call cs.OpenSQL(NavSQL)
                If cs.OK() Then
                    Do While cs.OK()
                        CategoryID = cs.GetInteger("CategoryID")
                        Call CS2.Open(ContentNameNewsletterStories, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & issueid & ")", "SortOrder,id")
                        If CS2.OK Then
                            CategoryName = cs.GetText("CategoryName")
                            If (CategoryName <> PreviousCategoryName) Then
                                AccessString = cn.GetCategoryAccessString(cp, cs.GetInteger("CategoryID"))
                                If AccessString <> "" Then
                                    repeatList &= "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                                End If
                                repeatList &= vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                                If AccessString <> "" Then
                                    repeatList &= "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                                End If
                                PreviousCategoryName = CategoryName
                            End If
                            '
                            Do While CS2.OK
                                '
                                Call repeatItem.Load(newsNavItem)
                                WorkingStoryId = CS2.GetInteger("ID")
                                AccessString = cn.GetArticleAccessString(cp, WorkingStoryId)
                                storyCaption = CS2.GetEditLink() & CS2.GetText("Name")
                                Call repeatItem.SetInner(".newsNavItemCaption", storyCaption)
                                If AccessString <> "" Then
                                    repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>")
                                End If
                                QS = cp.Doc.RefreshQueryString
                                QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, CStr(WorkingStoryId), True)
                                QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                                If AccessString <> "" Then
                                    repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                                End If
                                repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                                '
                                ArticleCount = ArticleCount + 1
                                Call CS2.GoNext()
                            Loop
                        End If
                        Call CS2.Close()
                        '
                        Call cs.GoNext()
                    Loop
                End If
                Call cs.Close()
                '
                Call cs.Open(ContentNameNewsletterStories, "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & issueid & ")", "SortOrder,DateAdded")
                If cs.OK() Then
                    If ArticleCount > 0 Then
                        '
                        ' This is a list of uncategorized articles following the categories -- give it a heading
                        '
                        CategoryName = cp.Site.GetText("Newsletter Nav Caption Other Articles", "Other Articles")
                        repeatList &= vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                    End If
                    Do While cs.OK()
                        Call repeatItem.Load(newsNavItem)
                        WorkingStoryId = cs.GetInteger("ID")
                        AccessString = cn.GetArticleAccessString(cp, WorkingStoryId)
                        storyCaption = cs.GetText("Name")
                        'storyCaption = CS.GetEditLink() & CS.GetText("Name")
                        If AccessString <> "" Then
                            repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>")
                        End If
                        Call repeatItem.SetInner(".newsNavItemCaption", storyCaption)
                        If cs.GetBoolean("AllowReadMore") Then
                            '
                            ' link to the story page
                            '
                            QS = cp.Doc.RefreshQueryString
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, CStr(WorkingStoryId), True)
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                        Else
                            '
                            ' link to the bookmark 'story#' on the cover
                            '
                            QS = "?" & cp.Doc.RefreshQueryString
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, "", False)
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormCover, True)
                            QS = QS & "#story" & WorkingStoryId
                        End If
                        If AccessString <> "" Then
                            repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                        End If
                        repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                        Call cs.GoNext()
                    Loop
                End If
                Call cs.Close()
                '
                ' Link to Current Issues
                '
                currentIssueId = cn.GetCurrentIssueID(cp, NewsletterID)
                If (issueid <> currentIssueId) And (currentIssueId <> 0) Then
                    Call repeatItem.Load(newsNavItem)
                    Call repeatItem.SetInner(".newsNavItemCaption", cp.Site.GetText(SitePropertyCurrentIssue, "Current Issue"))
                    repeatList &= repeatItem.GetHtml().Replace("?", "?" & cp.Doc.RefreshQueryString & RequestNameFormID & "=" & FormCover)
                End If
                '
                ' Display Archive Link if there are archive issues
                ' can not just lookup issues that are not the issueid because if you are editing a future issue, the current issue shows up as an archive
                '
                ThisSQL = "SELECT TOP 2 ID From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")"
                Call cs.OpenSQL(ThisSQL)
                If cs.OK Then
                    '
                    ' First one is the current issue
                    '
                    Call cs.GoNext()
                    If cs.OK Then
                        '
                        ' If there are more then one published issues, the others are archive issues
                        '
                        Call repeatItem.Load(newsNavItem)
                        Call repeatItem.SetInner(".newsNavItemCaption", cp.Site.GetText(SitePropertyIssueArchive, "Archives"))
                        repeatList &= repeatItem.GetHtml().Replace("?", "?" & cp.Doc.RefreshQueryString & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive)
                    End If
                End If
                Call cs.Close()
                '
                Call layout.SetInner(".newsNavList", repeatList)
                '
                returnHtml = layout.GetHtml()
            Catch ex As Exception

            End Try
            Return returnHtml
        End Function
        '
        Private Function GetArchiveLink(ByVal cp As CPBaseClass, ByVal newsletterId As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
            Stream &= "<a class=""caption"" href=""" & "?" & cp.Doc.RefreshQueryString & RequestNameNewsletterID & "=" & newsletterId & "&" & RequestNameFormID & "=" & FormArchive & """>" & cp.Site.GetText(SitePropertyIssueArchive, "Archives") & "</a>"
            'stream &=  "<a class=""caption"" href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & "?" & cp.Doc.RefreshQueryString & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & cp.site.getText(SitePropertyIssueArchive, "Archives", True) & "</a>"
            '
            GetArchiveLink = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("NavigationClass", "GetArchiveLink")
        End Function
        '
        Private Function GetCurrentIssueLink(ByVal cp As CPBaseClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
            Stream &= "<a class=""caption"" href=""" & "?" & cp.Doc.RefreshQueryString & RequestNameFormID & "=" & FormCover & """>" & cp.Site.GetText(SitePropertyCurrentIssue, "Current Issue") & "</a>"
            'stream &=  "<a class=""caption"" href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & "?" & cp.Doc.RefreshQueryString & RequestNameFormID & "=" & FormIssue & """>" & cp.site.getText(SitePropertyCurrentIssue, "Current Issue", True) & "</a>"
            '
            GetCurrentIssueLink = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("NavigationClass", "GetCurrentIssueLink")
        End Function
    End Class
End Namespace
