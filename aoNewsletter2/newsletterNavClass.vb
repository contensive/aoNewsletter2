
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
        Public Function GetContent(cp As CPBaseClass, isContentManager As Boolean, Optional LocalGroupID As Integer = 0) As String
            Dim returnString As String = ""
            Try

                Dim ErrorString As String
                Dim FormID As Integer
                Dim IssueID As Integer
                Dim IssuePageID As Integer
                Dim CurrentIssueID As Integer
                Dim NewsletterID As Integer
                Dim isManager As Boolean

                Dim NewsletterName As String
                Dim NewsletterProperty As String
                Dim Parts() As String
                Dim Stream As String
                Dim NavMode As String
                Dim OptionArray() As String
                Dim BracketPosition As Integer
                Dim cn As New newsletterCommonClass
                '
                NewsletterName = cp.Doc.GetText("Newsletter")
                If NewsletterName <> "" Then
                    '
                    ' If newsletterNavClass used without PageClass, Newsletter is in the OptionString, Issue is in QS
                    '
                    NewsletterID = cp.Content.GetRecordID(ContentNameNewsletters, NewsletterName)
                    Call cp.Site.TestPoint("GetIssueID call 3, NewsletterID=" & NewsletterID)
                    IssueID = cn.GetIssueID(cp, NewsletterID)
                    IssuePageID = cp.Doc.GetInteger(RequestNameIssuePageID)
                    FormID = cp.Doc.GetInteger(RequestNameFormID)
                Else
                    '
                    ' Without a Newsletter option, assume newsletterNavClass is used within a PageClass
                    ' Get the Issue and Newsletter from the visit properties set in PageClass
                    '
                    NewsletterProperty = cp.Visit.GetText(VisitPropertyNewsletter)
                    Parts = Split(NewsletterProperty, ".")
                    If UBound(Parts) > 2 Then
                        NewsletterID = cp.Utils.EncodeInteger(Parts(0))
                        IssueID = cp.Utils.EncodeInteger(Parts(1))
                        IssuePageID = cp.Utils.EncodeInteger(Parts(2))
                        FormID = cp.Utils.EncodeInteger(Parts(3))
                    End If
                End If
                CurrentIssueID = cn.GetCurrentIssueID(cp, NewsletterID)
                '
                '        IssueID = cn.GetIssueID(cp, NewsletterID)
                '
                Stream = GetNavigationVertical(cp, IssueID, CurrentIssueID, NewsletterID, isContentManager, FormID, LocalGroupID)
                '
                returnString = Stream
                '
            Catch ex As Exception

            End Try
            Return returnString
        End Function
        '
        Private Function GetNavigationVertical(cp As CPBaseClass, issueid As Integer, currentIssueId As Integer, NewsletterID As Integer, isContentManager As Boolean, FormID As Integer, Optional LocalGroupID As Integer = 0) As String
            Dim returnString As String = ""
            Try
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CSPointer As CPCSBaseClass = cp.CSNew()
                Dim ThisSQL As String
                Dim Controls As String
                Dim Link As String
                Dim Stream As String
                Dim WorkingIssuePageID As Integer
                Dim NavSQL As String
                Dim CategoryName As String
                Dim PreviousCategoryName As String
                Dim cn As New newsletterCommonClass
                Dim AccessString As String
                Dim CategoryID As Integer
                Dim QS As String
                Dim ArticleCount As Integer
                '
                Stream = "<div class=""NewsletterNav"">"
                Stream &= "<div class=""caption"">" & cp.Site.GetText(SitePropertyPageListCaption, "In This Issue") & "</div>"
                QS = cp.Doc.RefreshQueryString
                QS = cp.Utils.ModifyQueryString(QS, RequestNameIssueID, CStr(issueid), True)
                QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormIssue, True)
                Stream &= "<div class=""PageList""><a href=""?" & QS & """>Home</a></div>"
                '
                NavSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder, NIC.Name AS CategoryName"
                NavSQL = NavSQL & " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR"
                NavSQL = NavSQL & " Where (NIC.ID = NIR.CategoryID)"
                NavSQL = NavSQL & " AND (NIR.NewsletterIssueID=" & issueid & ")"
                NavSQL = NavSQL & " AND (NIC.Active<>0)"
                NavSQL = NavSQL & " AND (NIR.Active<>0)"
                NavSQL = NavSQL & " ORDER BY NIR.SortOrder"
                '
                Call CS.OpenSQL(NavSQL)
                If CS.OK() Then
                    Do While CS.OK()
                        CategoryID = CS.GetInteger("CategoryID")
                        Call CS2.Open(ContentNameNewsletterIssuePages, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & issueid & ")", "SortOrder")
                        If CS2.OK Then
                            CategoryName = CS.GetText("CategoryName")
                            If (CategoryName <> PreviousCategoryName) Then
                                AccessString = cn.GetCategoryAccessString(cp, CS.GetInteger("CategoryID"))
                                If AccessString <> "" Then
                                    Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                                End If
                                Stream &= vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                                If AccessString <> "" Then
                                    Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                                End If
                                PreviousCategoryName = CategoryName
                            End If
                            '
                            Do While CS2.OK
                                '
                                WorkingIssuePageID = CS2.GetInteger("ID")
                                AccessString = cn.GetArticleAccessString(cp, WorkingIssuePageID)
                                If AccessString <> "" Then
                                    Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                                End If
                                ' 1/1/09 - JK - always links to root page '/', removed path from link, added modify call incase requests are already in the qs
                                QS = cp.Doc.RefreshQueryString
                                QS = cp.Utils.ModifyQueryString(QS, RequestNameIssuePageID, CStr(WorkingIssuePageID), True)
                                QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                                Stream &= "<div class=""PageList"">" & CS2.GetEditLink() & "<a href=""?" & QS & """>" & CS.GetText("Name") & "</a></div>"
                                'stream &=  "<div class=""PageList"">" & cs.GetEditLink(CS2) & "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & "?" & cp.Doc.RefreshQueryString & RequestNameIssuePageID & "=" & WorkingIssuePageID & "&" & RequestNameFormID & "=" & FormDetails & """>" & cs.gettext2, "Name") & "</a></div>"
                                If AccessString <> "" Then
                                    Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                                End If
                                '
                                ArticleCount = ArticleCount + 1
                                Call CS2.GoNext()
                            Loop
                        End If
                        Call CS2.Close()
                        '
                        Call CS.GoNext()
                    Loop
                End If
                Call CS.Close()
                '
                Call CS.Open(ContentNameNewsletterIssuePages, "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & issueid & ")", "SortOrder,DateAdded")
                If CS.OK() Then
                    If ArticleCount > 0 Then
                        '
                        ' This is a list of uncategorized articles following the categories -- give it a heading
                        '
                        CategoryName = cp.Site.GetText("Newsletter Nav Caption Other Articles", "Other Articles")
                        Stream &= vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                    End If
                    Do While CS.OK()
                        WorkingIssuePageID = CS.GetInteger("ID")
                        AccessString = cn.GetArticleAccessString(cp, WorkingIssuePageID)
                        If AccessString <> "" Then
                            Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                        End If
                        ' 1/1/09 - JK - always links to root page '/', removed path from link, added modify call incase requests are already in the qs
                        If CS.GetBoolean("AllowReadMore") Then
                            '
                            ' link to the story page
                            '
                            QS = cp.Doc.RefreshQueryString
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameIssuePageID, CStr(WorkingIssuePageID), True)
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                            Stream &= "<div class=""PageList"">" & CS.GetEditLink() & "<a href=""?" & QS & """>" & CS.GetText("Name") & "</a></div>"
                        Else
                            '
                            ' link to the bookmark 'story#' on the cover
                            '
                            QS = "?" & cp.Doc.RefreshQueryString
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameIssuePageID, "", False)
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormIssue, True)
                            QS = QS & "#story" & WorkingIssuePageID
                            Stream &= "<div class=""PageList"">" & CS.GetEditLink() & "<a href=""" & QS & """>" & CS.GetText("Name") & "</a></div>"
                        End If
                        'stream &=  "<div class=""PageList"">" & cs.GetEditLink() & "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & "?" & cp.Doc.RefreshQueryString & RequestNameIssuePageID & "=" & WorkingIssuePageID & "&" & RequestNameFormID & "=" & FormDetails & """>" &cs.getText( "Name") & "</a></div>"
                        If AccessString <> "" Then
                            Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                        End If
                        Call CS.GoNext()
                    Loop
                End If
                Call CS.Close()
                '
                ' Link to Current Issues
                '
                If (issueid <> currentIssueId) And (currentIssueId <> 0) Then
                    Stream &= vbCrLf & "<div class=""caption""><a href=""" & "?" & cp.Doc.RefreshQueryString & RequestNameFormID & "=" & FormIssue & """>" & cp.Site.GetText(SitePropertyCurrentIssue, "Current Issue") & "</a></div>"
                End If
                '
                ' Display Archive Link if there are archive issues
                ' can not just lookup issues that are not the issueid because if you are editing a future issue, the current issue shows up as an archive
                '
                ThisSQL = "SELECT TOP 2 ID From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")"
                Call CS.OpenSQL(ThisSQL)
                If CS.OK Then
                    '
                    ' First one is the current issue
                    '
                    Call CS.GoNext()
                    If CS.OK Then
                        '
                        ' If there are more then one published issues, the others are archive issues
                        '
                        Stream &= vbCrLf & "<div class=""caption""><a href=""" & "?" & cp.Doc.RefreshQueryString & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & cp.Site.GetText(SitePropertyIssueArchive, "Archives") & "</a></div>"
                    End If
                End If
                Call CS.Close()
                '
                ' Admin Links
                '
                If (isContentManager) And (FormID <> FormEmail) Then
                    Controls = ""
                    If Link <> "" Then
                        Controls = Controls & vbCrLf & "<div class=""LinkLine"">" & Link & "</div>"
                    End If
                    'Controls = Controls & cn.GetAuthoringLinks(cp, IssuePageID, IssueID, NewsletterID, "?" & cp.Doc.RefreshQueryString)
                    If Controls <> "" Then
                        Stream &= "<BR /><BR />" & cn.GetAdminHintWrapper(cp, Controls)
                    End If
                End If
                '
                returnString = Stream

            Catch ex As Exception

            End Try
            Return returnString
        End Function
        '
        Private Function GetArchiveLink(cp As CPBaseClass, newsletterId As Integer) As String
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
        Private Function GetCurrentIssueLink(cp As CPBaseClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
            Stream &= "<a class=""caption"" href=""" & "?" & cp.Doc.RefreshQueryString & RequestNameFormID & "=" & FormIssue & """>" & cp.Site.GetText(SitePropertyCurrentIssue, "Current Issue") & "</a>"
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
