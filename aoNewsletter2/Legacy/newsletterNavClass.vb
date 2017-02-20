
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
        Friend Function GetNav(ByVal cp As CPBaseClass, ByVal issueid As Integer, ByVal NewsletterID As Integer, ByVal isContentManager As Boolean, ByVal FormID As Integer, ByVal newsNav As String, ByVal currentIssueId As Integer) As String
            Dim returnHtml As String = ""
            Try
                Dim layout As CPBlockBaseClass = cp.BlockNew()
                Dim repeatItem As CPBlockBaseClass = cp.BlockNew()
                '
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CSPointer As CPCSBaseClass = cp.CSNew()
                Dim ThisSQL As String
                'Dim WorkingStoryId As Integer
                Dim NavSQL As String
                Dim CategoryName As String
                Dim PreviousCategoryName As String = ""
                Dim cn As New newsletterCommonClass
                Dim AccessString As String
                Dim CategoryID As Integer
                Dim QS As String
                Dim ArticleCount As Integer
                Dim newsNavStoryItem As String
                Dim newsNavCategoryItem As String
                'Dim storyCaption As String
                Dim repeatList As String = ""
                '
                Call layout.Load(newsNav)
                newsNavStoryItem = layout.GetOuter(".newsNavStoryItem")
                newsNavCategoryItem = layout.GetOuter(".newsNavCategoryItem")
                '
                Call repeatItem.Load(newsNavStoryItem)
                QS = cp.Doc.RefreshQueryString
                QS = cp.Utils.ModifyQueryString(QS, RequestNameIssueID, CStr(issueid), True)
                QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormCover, True)
                Call repeatItem.SetInner(".newsNavItemCaption", "Home")
                'repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                repeatList &= repeatItem.GetHtml().Replace("href=""?""", "href=""?" & QS & """")
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
                                '
                                Call repeatItem.Load(newsNavCategoryItem)
                                Call repeatItem.SetInner(".newsNavItemCaption", CategoryName)
                                repeatList &= repeatItem.GetHtml()
                                '
                                If AccessString <> "" Then
                                    repeatList &= "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                                End If
                                PreviousCategoryName = CategoryName
                            End If
                            '
                            Do While CS2.OK
                                '
                                repeatList &= getNavItem(cp, cn, CS2, newsNavStoryItem)
                                'Call repeatItem.Load(newsNavStoryItem)
                                'WorkingStoryId = CS2.GetInteger("ID")
                                'AccessString = cn.GetArticleAccessString(cp, WorkingStoryId)
                                'storyCaption = CS2.GetText("Name")
                                'Call repeatItem.SetInner(".newsNavItemCaption", storyCaption)
                                'If AccessString <> "" Then
                                '    repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>")
                                'End If
                                'QS = cp.Doc.RefreshQueryString
                                'QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, CStr(WorkingStoryId), True)
                                'QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                                'If AccessString <> "" Then
                                '    repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                                'End If
                                'repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
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
                        Call repeatItem.Load(newsNavCategoryItem)
                        Call repeatItem.SetInner(".newsNavItemCaption", CategoryName)
                        repeatList &= repeatItem.GetHtml()
                    End If
                    Do While cs.OK()
                        repeatList &= getNavItem(cp, cn, cs, newsNavStoryItem)
                        'Call repeatItem.Load(newsNavStoryItem)
                        'WorkingStoryId = cs.GetInteger("ID")
                        'AccessString = cn.GetArticleAccessString(cp, WorkingStoryId)
                        'storyCaption = cs.GetText("Name")
                        ''storyCaption = CS.GetEditLink() & CS.GetText("Name")
                        'If AccessString <> "" Then
                        '    repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>")
                        'End If
                        'Call repeatItem.SetInner(".newsNavItemCaption", storyCaption)
                        'If Not cn.isBlank(cp, cs.GetText("body")) Then
                        '    'If cs.GetBoolean("AllowReadMore") Then
                        '    '
                        '    ' link to the story page
                        '    '
                        '    QS = cp.Doc.RefreshQueryString
                        '    QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, CStr(WorkingStoryId), True)
                        '    QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                        'Else
                        '    '
                        '    ' link to the bookmark 'story#' on the cover
                        '    '
                        '    QS = "?" & cp.Doc.RefreshQueryString
                        '    QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, "", False)
                        '    QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormCover, True)
                        '    QS = QS & "#story" & WorkingStoryId
                        'End If
                        'If AccessString <> "" Then
                        '    repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                        'End If
                        'repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                        Call cs.GoNext()
                    Loop
                End If
                Call cs.Close()
                '
                ' Link to Current Issues
                '
                If (issueid <> currentIssueId) And (currentIssueId <> 0) Then
                    QS = cp.Doc.RefreshQueryString
                    QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormCover)
                    Call repeatItem.Load(newsNavStoryItem)
                    Call repeatItem.SetInner(".newsNavItemCaption", cp.Site.GetText(SitePropertyCurrentIssue, "Current Issue"))
                    'repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                    repeatList &= repeatItem.GetHtml().Replace("href=""?""", "href=""?" & QS & """")
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
                        Call repeatItem.Load(newsNavStoryItem)
                        Call repeatItem.SetInner(".newsNavItemCaption", cp.Site.GetText(SitePropertyIssueArchive, "Archives"))
                        QS = cp.Doc.RefreshQueryString
                        'QS = cp.Utils.ModifyQueryString(QS, RequestNameNewsletterID, NewsletterID)
                        QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormArchive)
                        'repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                        repeatList &= repeatItem.GetHtml().Replace("href=""?""", "href=""?" & QS & """")
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
            Dim Stream As String = ""
            Dim qs As String = ""
            '
            qs = cp.Doc.RefreshQueryString
            qs = cp.Utils.ModifyQueryString(qs, RequestNameFormID, FormArchive)
            Stream &= "<a class=""caption"" href=""?" & qs & """>" & cp.Site.GetText(SitePropertyIssueArchive, "Archives") & "</a>"
            GetArchiveLink = Stream
        End Function
        '
        Private Function GetCurrentIssueLink(ByVal cp As CPBaseClass) As String
            Dim Stream As String = ""
            Dim qs As String = ""
            '
            qs = cp.Doc.RefreshQueryString
            qs = cp.Utils.ModifyQueryString(qs, RequestNameFormID, FormCover)
            Stream &= "<a class=""caption"" href=""?" & qs & """>" & cp.Site.GetText(SitePropertyCurrentIssue, "Current Issue") & "</a>"
            GetCurrentIssueLink = Stream
        End Function
        '
        Private Function getNavItem(ByVal cp As CPBaseClass, ByVal cn As newsletterCommonClass, ByVal cs As CPCSBaseClass, ByVal newsNavStoryItemLayout As String) As String
            Dim returnHtml As String = ""
            Try
                Dim repeatItem As CPBlockBaseClass = cp.BlockNew()
                Dim WorkingStoryId As Integer
                Dim accessString As String
                Dim storyCaption As String
                Dim qs As String
                '
                Call repeatItem.Load(newsNavStoryItemLayout)
                WorkingStoryId = cs.GetInteger("ID")
                accessString = cn.GetArticleAccessString(cp, WorkingStoryId)
                storyCaption = cs.GetText("Name")
                'storyCaption = CS.GetEditLink() & CS.GetText("Name")
                If accessString <> "" Then
                    repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & accessString & """>")
                End If
                Call repeatItem.SetInner(".newsNavItemCaption", storyCaption)
                If Not cn.isBlank(cp, cs.GetText("body")) Then
                    'If cs.GetBoolean("AllowReadMore") Then
                    '
                    ' link to the story page
                    '
                    qs = cp.Doc.RefreshQueryString
                    qs = cp.Utils.ModifyQueryString(qs, RequestNameStoryId, CStr(WorkingStoryId), True)
                    qs = cp.Utils.ModifyQueryString(qs, RequestNameFormID, FormDetails, True)
                Else
                    '
                    ' link to the bookmark 'story#' on the cover
                    '
                    qs = cp.Doc.RefreshQueryString
                    qs = cp.Utils.ModifyQueryString(qs, RequestNameStoryId, "", False)
                    qs = cp.Utils.ModifyQueryString(qs, RequestNameFormID, FormCover, True)
                    qs = qs & "#story" & WorkingStoryId
                End If
                If accessString <> "" Then
                    repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                End If
                'returnHtml = repeatItem.GetHtml().Replace("?", "?" & qs)
                returnHtml = repeatItem.GetHtml().Replace("href=""?""", "href=""?" & qs & """")
            Catch ex As Exception
                Call handleError(cp, ex, "getNavItem")
            End Try
            Return returnHtml
        End Function
        '
    End Class
End Namespace
