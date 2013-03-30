
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterNavClass
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
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterNavClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Private WorkingQueryStringPlus As String
        Private ErrorString As String
        '
        Private FormID As Long
        Private IssueID As Long
        Private IssuePageID As Long
        '
        ' The issueid of the most current issue -- may not be this issue
        '
        Private CurrentIssueID As Long
        '
        Private NewsletterID As Long

        Private isManager As Boolean

        Private Main As Object
        Private Csv As Object

        Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
            On Error GoTo ErrorTrap

            Csv = CsvObject

            Call Init(MainObject)

            Execute = GetContent(OptionString)

            Exit Function
ErrorTrap:
            Call HandleError("NavClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
        End Function

        Public Sub Init(MainObject As Object)
            '
            Main = MainObject
            '
            Dim Common As New CommonClass
            '
            Call Common.UpgradeAddOn(Main)

            isManager = Main.IsContentManager("Newsletters")

            Exit Sub
            '
ErrorTrap:
            Call HandleError("NavigationClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
        End Sub
        '
Public Function GetContent(OptionString As String, Optional LocalGroupID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim NewsletterName As String
            Dim NewsletterProperty As String
            Dim Parts() As String
            Dim Stream As String
            Dim NavMode As String
            Dim OptionArray() As String
            Dim BracketPosition As Long
            Dim Common As New CommonClass
            '
            If Not (Main Is Nothing) Then
                '
                NewsletterName = Main.GetAddonOption("Newsletter", OptionString)
                If NewsletterName <> "" Then
                    '
                    ' If NavClass used without PageClass, Newsletter is in the OptionString, Issue is in QS
                    '
                    NewsletterID = Main.GetRecordID(ContentNameNewsletters, NewsletterName)
                    Call Main.TestPoint("GetIssueID call 3, NewsletterID=" & NewsletterID)
                    IssueID = Common.GetIssueID(Main, NewsletterID)
                    IssuePageID = Main.GetStreamInteger(RequestNameIssuePageID)
                    FormID = Main.GetStreamInteger(RequestNameFormID)
                Else
                    '
                    ' Without a Newsletter option, assume NavClass is used within a PageClass
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
                CurrentIssueID = Common.GetCurrentIssueID(Main, NewsletterID)
                '
                WorkingQueryStringPlus = Main.RefreshQueryString
                '
                If WorkingQueryStringPlus = "" Then
                    WorkingQueryStringPlus = "?"
                Else
                    WorkingQueryStringPlus = "?" & WorkingQueryStringPlus & "&"
                End If
                '
                '        IssueID = Common.GetIssueID(Main, NewsletterID)
                '
                Stream = GetNavigationVertical(LocalGroupID)
                '
                GetContent = Stream
                '
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("NavigationClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
Private Function GetNavigationVertical(Optional LocalGroupID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim CSPointer As Long
            Dim ThisSQL As String
            Dim Controls As String
            Dim Link As String
            Dim Stream As String
            Dim CS As String
            Dim WorkingIssuePageID As Long
            Dim NavSQL As String
            Dim CategoryName As String
            Dim PreviousCategoryName As String
            Dim Common As New CommonClass
            Dim AccessString As String
            Dim CS2 As Long
            Dim CategoryID As Long
            Dim QS As String
            Dim ArticleCount As Long
            '
            Stream = "<div class=""NewsletterNav"">"
            Stream = Stream & "<div class=""caption"">" & Main.GetSiteProperty(SitePropertyPageListCaption, "In This Issue", True) & "</div>"
            QS = WorkingQueryStringPlus
            QS = ModifyQueryString(QS, RequestNameIssueID, CStr(IssueID), True)
            QS = ModifyQueryString(QS, RequestNameFormID, FormIssue, True)
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
                    CategoryID = Main.GetCSInteger(CS, "CategoryID")
                    CS2 = Main.OpenCSContent(ContentNameNewsletterIssuePages, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & IssueID & ")", "SortOrder")
                    If Main.IsCSOK(CS2) Then
                        CategoryName = Main.GetCSText(CS, "CategoryName")
                        If (CategoryName <> PreviousCategoryName) Then
                            AccessString = Common.GetCategoryAccessString(Main, Main.GetCSInteger(CS, "CategoryID"))
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
                            WorkingIssuePageID = Main.GetCSInteger(CS2, "ID")
                            AccessString = Common.GetArticleAccessString(Main, WorkingIssuePageID)
                            If AccessString <> "" Then
                                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                            End If
                            ' 1/1/09 - JK - always links to root page '/', removed path from link, added modify call incase requests are already in the qs
                            QS = WorkingQueryStringPlus
                            QS = ModifyQueryString(QS, RequestNameIssuePageID, CStr(WorkingIssuePageID), True)
                            QS = ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                            Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS2) & "<a href=""" & QS & """>" & Main.GetCSText(CS2, "Name") & "</a></div>"
                            'Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS2) & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & WorkingIssuePageID & "&" & RequestNameFormID & "=" & FormDetails & """>" & Main.GetCSText(CS2, "Name") & "</a></div>"
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
            Call Main.CloseCS(CS)
            '
            CS = Main.OpenCSContent(ContentNameNewsletterIssuePages, "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")", "SortOrder,DateAdded")
            If Main.CSOK(CS) Then
                If ArticleCount > 0 Then
                    '
                    ' This is a list of uncategorized articles following the categories -- give it a heading
                    '
                    CategoryName = Main.GetSiteProperty("Newsletter Nav Caption Other Articles", "Other Articles")
                    Stream = Stream & vbCrLf & "<div class=""NewsletterNavTopic"">" & CategoryName & "</div>"
                End If
                Do While Main.CSOK(CS)
                    WorkingIssuePageID = Main.GetCSInteger(CS, "ID")
                    AccessString = Common.GetArticleAccessString(Main, WorkingIssuePageID)
                    If AccessString <> "" Then
                        Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>"
                    End If
                    ' 1/1/09 - JK - always links to root page '/', removed path from link, added modify call incase requests are already in the qs
                    If Main.GetCSBoolean(CS, "AllowReadMore") Then
                        '
                        ' link to the story page
                        '
                        QS = WorkingQueryStringPlus
                        QS = ModifyQueryString(QS, RequestNameIssuePageID, CStr(WorkingIssuePageID), True)
                        QS = ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                        Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS) & "<a href=""" & QS & """>" & Main.GetCSText(CS, "Name") & "</a></div>"
                    Else
                        '
                        ' link to the bookmark 'story#' on the cover
                        '
                        QS = WorkingQueryStringPlus
                        QS = ModifyQueryString(QS, RequestNameIssuePageID, "", False)
                        QS = ModifyQueryString(QS, RequestNameFormID, FormIssue, True)
                        QS = QS & "#story" & WorkingIssuePageID
                        Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS) & "<a href=""" & QS & """>" & Main.GetCSText(CS, "Name") & "</a></div>"
                    End If
                    'Stream = Stream & "<div class=""PageList"">" & Main.GetCSRecordEditLink(CS) & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & WorkingIssuePageID & "&" & RequestNameFormID & "=" & FormDetails & """>" & Main.GetCSText(CS, "Name") & "</a></div>"
                    If AccessString <> "" Then
                        Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
                    End If
                    Call Main.NextCSRecord(CS)
                Loop
            End If
            Call Main.CloseCS(CS)
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
                'Controls = Controls & Common.GetAuthoringLinks(Main, IssuePageID, IssueID, NewsletterID, WorkingQueryStringPlus)
                If Controls <> "" Then
                    Stream = Stream & "<BR /><BR />" & Main.GetAdminHintWrapper(Controls)
                End If
            End If
            '
            GetNavigationVertical = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("NavigationClass", "GetNavigationVertical", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetArchiveLink() As String
            On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
            Stream = Stream & "<a class=""caption"" href=""" & WorkingQueryStringPlus & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & Main.GetSiteProperty(SitePropertyIssueArchive, "Archives", True) & "</a>"
            'Stream = Stream & "<a class=""caption"" href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameNewsletterID & "=" & NewsletterID & "&" & RequestNameFormID & "=" & FormArchive & """>" & Main.GetSiteProperty(SitePropertyIssueArchive, "Archives", True) & "</a>"
            '
            GetArchiveLink = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("NavigationClass", "GetArchiveLink", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetCurrentIssueLink() As String
            On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
            Stream = Stream & "<a class=""caption"" href=""" & WorkingQueryStringPlus & RequestNameFormID & "=" & FormIssue & """>" & Main.GetSiteProperty(SitePropertyCurrentIssue, "Current Issue", True) & "</a>"
            'Stream = Stream & "<a class=""caption"" href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameFormID & "=" & FormIssue & """>" & Main.GetSiteProperty(SitePropertyCurrentIssue, "Current Issue", True) & "</a>"
            '
            GetCurrentIssueLink = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("NavigationClass", "GetCurrentIssueLink", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        ' 1/1/09 JK - moved to common GetAuthoringLinks
        '
        'Private Function GetEmailLink() As String
        '    On Error GoTo ErrorTrap
        '    '
        '    Dim Stream As String
        '    '
        '    ' 1/1/09 JK - moved to common GetAuthoringLinks
        '    'If IssueID <> 0 Then
        '    '    ' 1/1/09 - JK - fixed link - always pointed to the current page in the site's root directory (/index.asp), should point to the current page
        '    '    Stream = Stream & "<a class=""caption"" href=""" & WorkingQueryStringPlus & RequestNameFormID & "=" & FormEmail & """>Create&nbsp;Email&nbsp;Version</a>"
        '    '    'Stream = Stream & "<a class=""caption"" href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameFormID & "=" & FormEmail & """>Create&nbsp;Email&nbsp;Version</a>"
        '    'End If
        '    '
        '    GetEmailLink = Stream
        '    '
        '    Exit Function
        'ErrorTrap:
        '    Call HandleError("NavigationClass", "GetEmailLink", Err.Number, Err.Source, Err.Description, True, False)
        '    End Function
        '
        ' 1/1/9 JK - moved to common admin links
        '
        'Private Function GetFutureIssues() As String
        '    On Error GoTo ErrorTrap
        '    '
        '    Dim QS As String
        '    Dim Stream As String
        '    Dim CSPointer As String
        '    '
        '    CSPointer = Main.OpenCSContent(ContentNameNewsletterIssues, "(PublishDate>" & Main.EncodeSQLDate(Now()) & ") OR (PublishDate is Null) OR (PublishDate=" & KmaEncodeSQLDate(0) & ")", "PublishDate desc")
        '    If Main.CSOK(CSPointer) Then
        '        Do While Main.CSOK(CSPointer)
        '            ' 1/1/09 - JK - always linked to root path, also added qs incase request name was already in wqsp
        '            QS = WorkingQueryStringPlus
        '            QS = ModifyQueryString(QS, RequestNameIssueID, Main.GetCSInteger(CSPointer, "ID"), True)
        '            Stream = Stream & "<div class=""PageList""><a href=""" & QS & """>" & Main.GetCSText(CSPointer, "Name") & "</a></div>"
        '            'Stream = Stream & "<div class=""PageList""><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssueID & "=" & Main.GetCSInteger(CSPointer, "ID") & """>" & Main.GetCSText(CSPointer, "Name") & "</a></div>"
        '            Call Main.NextCSRecord(CSPointer)
        '        Loop
        '    End If
        '    Call Main.CloseCS(CSPointer)
        '    '
        '    GetFutureIssues = Stream
        '    '
        '    Exit Function
        'ErrorTrap:
        '    Call HandleError("NavigationClass", "GetFutureIssues", Err.Number, Err.Source, Err.Description, True, False)
        '    End Function

    End Class
End Namespace
