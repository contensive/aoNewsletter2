
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    '
    Public Class newsletterNavClass
        Inherits AddonBaseClass
        '
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
Public Function GetContent(cp As CPBaseClass, OptionString As String, Optional LocalGroupID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim WorkingQueryStringPlus As String
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
            Dim Common As New newsletterCommonClass
            '
            If True Then
                '
                NewsletterName = cp.Doc.GetText("Newsletter", OptionString)
                If NewsletterName <> "" Then
                    '
                    ' If newsletterNavClass used without PageClass, Newsletter is in the OptionString, Issue is in QS
                    '
                    NewsletterID = cp.Content.GetRecordID(ContentNameNewsletters, NewsletterName)
                    Call cp.Site.TestPoint("GetIssueID call 3, NewsletterID=" & NewsletterID)
                    IssueID = Common.GetIssueID(cp, NewsletterID)
                    IssuePageID = cp.Doc.GetInteger(RequestNameIssuePageID)
                    FormID = cp.Doc.GetInteger(RequestNameFormID)
                Else
                    '
                    ' Without a Newsletter option, assume newsletterNavClass is used within a PageClass
                    ' Get the Issue and Newsletter from the visit properties set in PageClass
                    '
                    NewsletterProperty = Main.GetVisitProperty(VisitPropertyNewsletter)
                    Parts() = Split(NewsletterProperty, ".")
                    If UBound(Parts) > 2 Then
                        NewsletterID = kmaEncodeInteger(Parts(0))
                        IssueID = kmaEncodeInteger(Parts(1))
                        IssuePageID = kmaEncodeInteger(Parts(2))
                        FormID = kmaEncodeInteger(Parts(3))
                    End If
                End If
                CurrentIssueID = Common.GetCurrentIssueID(cp, NewsletterID)
                '
                WorkingQueryStringPlus = cp.Doc.RefreshQueryString
                '
                If WorkingQueryStringPlus = "" Then
                    WorkingQueryStringPlus = "?"
                Else
                    WorkingQueryStringPlus = "?" & WorkingQueryStringPlus & "&"
                End If
                '
                '        IssueID = Common.GetIssueID(cp, NewsletterID)
                '
                Stream = GetNavigationVertical(LocalGroupID)
                '
                GetContent = Stream
                '
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("NavigationClass", "GetContent")
        End Function
        '
Private Function GetNavigationVertical(cp As CPBaseClass, Optional LocalGroupID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim CSPointer As CPCSBaseClass = cp.CSNew()
            Dim ThisSQL As String
            Dim Controls As String
            Dim Link As String
            Dim Stream As String
            Dim CS As String
            Dim WorkingIssuePageID As Integer
            Dim NavSQL As String
            Dim CategoryName As String
            Dim PreviousCategoryName As String
            Dim Common As New newsletterCommonClass
            Dim AccessString As String
            Dim CS2 As Integer
            Dim CategoryID As Integer
            Dim QS As String
            Dim ArticleCount As Integer
            '
            Stream = "<div class=""NewsletterNav"">"
            Stream = Stream & "<div class=""caption"">" & Main.GetSiteProperty(SitePropertyPageListCaption, "In This Issue", True) & "</div>"
            QS = WorkingQueryStringPlus
            QS = cp.Utils.ModifyQueryString(QS, RequestNameIssueID, CStr(IssueID), True)
            QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormIssue, True)
            Stream = Stream & "<div class=""PageList""><a href=""" & QS & """>Home</a></div>"
            '
            NavSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder, NIC.Name AS CategoryName"
            NavSQL = NavSQL & " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR"
            NavSQL = NavSQL & " Where (NIC.ID = NIR.CategoryID)"
            NavSQL = NavSQL & " AND (NIR.NewsletterIssueID=" & IssueID & ")"
            NavSQL = NavSQL & " AND (NIC.Active<>0)"
            NavSQL = NavSQL & " AND (NIR.Active<>0)"
            NavSQL = NavSQL & " ORDER BY NIR.SortOrder"
            '
            CS = Main.OpenCSSQL("Default", NavSQL)
            If Main.CSOK(CS) Then
                Do While Main.CSOK(CS)
                    CategoryID = CS.getInteger("CategoryID")
                    CS2 = Main.OpenCSContent(ContentNameNewsletterIssuePages, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & IssueID & ")", "SortOrder")
                    If Main.IsCSOK(CS2) Then
                        CategoryName = CS.getText("CategoryName")
                        If (CategoryName <> PreviousCategoryName) Then
                            AccessString = Common.GetCategoryAccessString(cp, CS.getInteger("CategoryID"))
                            If AccessString <> "" Then
                                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                            End If
                            Stream = Stream & vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                            If AccessString <> "" Then
                                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                            End If
                            PreviousCategoryName = CategoryName
                        End If
                        '
                        Do While Main.CSOK(CS2)
                            '
                            WorkingIssuePageID = CS.getInteger(CS2, "ID")
                            AccessString = Common.GetArticleAccessString(cp, WorkingIssuePageID)
                            If AccessString <> "" Then
                                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                            End If
                            ' 1/1/09 - JK - always links to root page '/', removed path from link, added modify call incase requests are already in the qs
                            QS = WorkingQueryStringPlus
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameIssuePageID, CStr(WorkingIssuePageID), True)
                            QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                            Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS2) & "<a href=""" & QS & """>" & Main.GetCSText(CS2, "Name") & "</a></div>"
                            'Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS2) & "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & WorkingIssuePageID & "&" & RequestNameFormID & "=" & FormDetails & """>" & Main.GetCSText(CS2, "Name") & "</a></div>"
                            If AccessString <> "" Then
                                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                            End If
                            '
                            ArticleCount = ArticleCount + 1
                            Call Main.NextCSRecord(CS2)
                        Loop
                    End If
                    Call Main.CloseCS(CS2)
                    '
                    Call Main.NextCSRecord(CS)
                Loop
            End If
            Call CS.close()
            '
            Call CS.open(ContentNameNewsletterIssuePages, "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")", "SortOrder,DateAdded")
            If Main.CSOK(CS) Then
                If ArticleCount > 0 Then
                    '
                    ' This is a list of uncategorized articles following the categories -- give it a heading
                    '
                    CategoryName = Main.GetSiteProperty("Newsletter Nav Caption Other Articles", "Other Articles")
                    Stream = Stream & vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                End If
                Do While Main.CSOK(CS)
                    WorkingIssuePageID = CS.getInteger("ID")
                    AccessString = Common.GetArticleAccessString(cp, WorkingIssuePageID)
                    If AccessString <> "" Then
                        Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                    End If
                    ' 1/1/09 - JK - always links to root page '/', removed path from link, added modify call incase requests are already in the qs
                    If Main.GetCSBoolean(CS, "AllowReadMore") Then
                        '
                        ' link to the story page
                        '
                        QS = WorkingQueryStringPlus
                        QS = cp.Utils.ModifyQueryString(QS, RequestNameIssuePageID, CStr(WorkingIssuePageID), True)
                        QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                        Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS) & "<a href=""" & QS & """>" & CS.getText("Name") & "</a></div>"
                    Else
                        '
                        ' link to the bookmark 'story#' on the cover
                        '
                        QS = WorkingQueryStringPlus
                        QS = cp.Utils.ModifyQueryString(QS, RequestNameIssuePageID, "", False)
                        QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormIssue, True)
                        QS = QS & "#story" & WorkingIssuePageID
                        Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS) & "<a href=""" & QS & """>" & CS.getText("Name") & "</a></div>"
                    End If
                    'Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS) & "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & WorkingIssuePageID & "&" & RequestNameFormID & "=" & FormDetails & """>" &cs.getText( "Name") & "</a></div>"
                    If AccessString <> "" Then
                        Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                    End If
                    Call Main.NextCSRecord(CS)
                Loop
            End If
            Call CS.close()
            '
            ' Link to Current Issues
            '
            If (IssueID <> CurrentIssueID) And (CurrentIssueID <> 0) Then
                Stream = Stream & vbCrLf & "<div class=""caption""><a href=""" & WorkingQueryStringPlus & RequestNameFormID & "=" & FormIssue & """>" & Main.GetSiteProperty(SitePropertyCurrentIssue, "Current Issue", True) & "</a></div>"
            End If
            '
            ' Display Archive Link if there are archive issues
            ' can not just lookup issues that are not the issueid because if you are editing a future issue, the current issue shows up as an archive
            '
            ThisSQL = "SELECT TOP 2 ID From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (NewsletterID=" & Main.EncodeSQLNumber(NewsletterID) & ")"
            CSPointer = Main.OpenCSSQL("Default", ThisSQL)
            If Main.CSOK(CSPointer) Then
                '
                ' First one is the current issue
                '
                Call Main.NextCSRecord(CSPointer)
                If Main.IsCSOK(CSPointer) Then
                    '
                    ' If there are more then one published issues, the others are archive issues
                    '
                    Stream = Stream & vbCrLf & "<div class=""caption""><a href=""" & WorkingQueryStringPlus & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & Main.GetSiteProperty(SitePropertyIssueArchive, "Archives", True) & "</a></div>"
                End If
            End If
            Call Main.CloseCS(CSPointer)
            '
            ' Admin Links
            '
            If (isManager) And (FormID <> FormEmail) Then
                Controls = ""
                If Link <> "" Then
                    Controls = Controls & vbCrLf & "<div class=""LinkLine"">" & Link & "</div>"
                End If
                'Controls = Controls & Common.GetAuthoringLinks(cp, IssuePageID, IssueID, NewsletterID, WorkingQueryStringPlus)
                If Controls <> "" Then
                    Stream = Stream & "<BR /><BR />" & Main.GetAdminHintWrapper(Controls)
                End If
            End If
            '
            GetNavigationVertical = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("NavigationClass", "GetNavigationVertical")
        End Function
        '
        Private Function GetArchiveLink(cp As CPBaseClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
            Stream = Stream & "<a class=""caption"" href=""" & WorkingQueryStringPlus & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & Main.GetSiteProperty(SitePropertyIssueArchive, "Archives", True) & "</a>"
            'Stream = Stream & "<a class=""caption"" href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & Main.GetSiteProperty(SitePropertyIssueArchive, "Archives", True) & "</a>"
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
            Stream = Stream & "<a class=""caption"" href=""" & WorkingQueryStringPlus & RequestNameFormID & "=" & FormIssue & """>" & Main.GetSiteProperty(SitePropertyCurrentIssue, "Current Issue", True) & "</a>"
            'Stream = Stream & "<a class=""caption"" href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameFormID & "=" & FormIssue & """>" & Main.GetSiteProperty(SitePropertyCurrentIssue, "Current Issue", True) & "</a>"
            '
            GetCurrentIssueLink = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("NavigationClass", "GetCurrentIssueLink")
        End Function
    End Class
End Namespace
