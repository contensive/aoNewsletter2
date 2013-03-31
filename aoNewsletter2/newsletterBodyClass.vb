Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterBodyClass
        '
        Private WorkingQueryStringPlus As String
        Private ErrorString As String
        '
        Private PageNumber As Integer
        '
        Private FormID As Integer
        Private RecordsPerPage As Integer
        '
        Private IssueID As Integer
        Private IssuePageID As Integer
        Private MonthSelected As Integer
        Private YearSelected As Integer
        Private ButtonValue As String
        Private RecordTop As Integer
        Private SearchKeywords As String

        Private isManager As Boolean

        Private NewsletterID As Integer
        Private archiveIssuesToDisplay As Integer
        '
        'Private Main As MainClass
        Private cp As CPBaseClass
        Private EncodeCopyNeeded As Boolean
        'Private Csv As Object
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub handleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterBodyClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Public Function GetContent(cp As CPBaseClass, OptionString As String) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            Dim cn As New newsletterCommonClass
            Dim NewsletterName As String
            Dim NewsletterProperty As String
            Dim Parts() As String
            '
            If True Then
                '
                NewsletterName = cp.Doc.GetText("Newsletter")
                archiveIssuesToDisplay = cp.Doc.GetInteger("Archive Issues To Display")
                '
                If NewsletterName <> "" Then
                    '
                    ' If newsletterNavClass used without PageClass, Newsletter is in the OptionString, Issue is in QS
                    '
                    NewsletterID = cp.Content.GetRecordID(ContentNameNewsletters, NewsletterName)
                    Call cp.Site.TestPoint("GetIssueID call 2, NewsletterID=" & NewsletterID)
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
                '
                MonthSelected = cp.Doc.GetInteger(RequestNameMonthSelectd)
                YearSelected = cp.Doc.GetInteger(RequestNameYearSelected)
                ButtonValue = cp.Doc.GetText("Button")
                SearchKeywords = cp.Doc.GetText(RequestNameSearchKeywords)
                '
                RecordsPerPage = cp.Site.GetInteger("Newsletter Search Results Records Per Page", "3")
                '
                RecordTop = cp.Doc.GetInteger(RequestNameRecordTop)
                '
                PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
                If PageNumber = 0 Then
                    PageNumber = 1
                End If

                Stream = GetForm(cp, cn)
                GetContent = Stream
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetContent")
        End Function
        '
        Private Function GetForm(cp As CPBaseClass, cn As newsletterCommonClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            ' Process forms
            '
            Select Case FormID
                Case FormArchive
                    Select Case ButtonValue
                        Case FormButtonViewNewsLetter
                            '
                            ' Archive form pressing the view button
                            '
                            FormID = FormIssue
                    End Select
            End Select
            '
            ' Display Forms
            '
            Select Case FormID
                Case FormArchive
                    Stream &= GetArchiveList(cp)
                Case FormDetails
                    Call cp.Site.TestPoint("GetForm Entering GetNewsletterBodyDetails")
                    Stream &= GetNewsletterBodyDetails(cp, cn, IssuePageID)
                Case Else
                    Call cp.Site.TestPoint("GetForm Entering GetNewsletterBodyOverview")
                    FormID = FormIssue
                    Stream &= GetNewsletterBodyOverview(cp, IssueID, IssuePageID, )
            End Select
            '    '
            '    Select Case ButtonValue
            '        Case FormButtonViewArchives
            ' '           stream &=  GetArchiveList()
            '        Case FormButtonViewNewsLetter
            ' '           stream &=  GetArchiveList()
            '    End Select
            '
            GetForm = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetForm")
        End Function
        '
        Private Function GetArchiveList(cp As CPBaseClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            '
            Dim Stream As String
            Dim Colors As String
            Dim ThisSQL As String
            Dim ThisSQL2 As String
            Dim MonthString As String
            Dim YearString As String
            Dim MonthCounter As Integer
            Dim YearCounter As Integer
            '
            Dim SelectedIssuePointer As Integer
            Dim SelectedIssue As String
            '
            Dim SearchResult As String
            Dim SearchSQL As String
            '
            Dim NumberofPages As Integer
            Dim PageCount As Integer
            '
            Dim sql2 As String
            Dim FileCount As Integer
            '
            Dim BriefCopyFileName As String
            '
            Dim RowCount As Integer
            Dim SQLCriteria As String
            '
            Dim YearsWanted As Integer
            Dim BlockSearchForm As Boolean
            Dim qs As String
            '
            YearsWanted = cp.Utils.EncodeInteger(cp.Site.GetText("Newsletter years wanted", 1))
            If YearsWanted < 1 Then
                YearsWanted = 1
            End If
            '
            If archiveIssuesToDisplay = 0 Then
                archiveIssuesToDisplay = 6
            End If
            '
            ' Get Total Archive count
            '
            PageCount = 1
            sql2 = " select count(nlp.id) as count"
            sql2 = sql2 & " from newsletterissues nl, newsletterissuepages nlp"
            sql2 = sql2 & " Where (NL.ID = nlp.newsletterid)"
            sql2 = sql2 & " AND (NL.NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")"
            If MonthSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & MonthSelected
            End If
            If YearSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & YearSelected
            End If
            If SearchKeywords <> "" Then
                sql2 = sql2 & " and ((nlp.Body like '%" & SearchKeywords & "%' )or (nlp.name  like '%" & SearchKeywords & "%') or (nlp.Overview  like '%" & SearchKeywords & "%'))"
            End If
            If cs.OpenSQL(sql2) Then
                FileCount = cs.GetInteger("count")
                NumberofPages = FileCount / RecordsPerPage
                If NumberofPages <> Int(NumberofPages) Then
                    NumberofPages = NumberofPages + 1
                    NumberofPages = Int(NumberofPages)
                End If
                If NumberofPages = 0 Then
                    NumberofPages = 1
                End If
            End If
            Call cs.Close()
            '
            'Colors = "#ffffff"
            '
            '
            If (ButtonValue <> FormButtonViewNewsLetter) And (ButtonValue <> FormButtonViewArchives) Then
                '
                ' List a page of archive issues
                '
                If (MonthSelected = 0) And (YearSelected = 0) Then
                    'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                    '
                    'ThisSQL = " SELECT  TOP 6 * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & cp.db.encodesqlNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    ThisSQL = " SELECT  TOP " & archiveIssuesToDisplay & " * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    '
                    Call cs.OpenSQL(ThisSQL)
                    If cs.OK Then
                        Stream &= cp.Content.GetCopy(PageNameArchives, "<h2>Archive Issues</h2>")
                        Do While cs.OK
                            Stream &= cs.GetEditLink() & "<a href=""" & "?" & cp.Doc.RefreshQueryString & RequestNameIssueID & "=" & cs.GetInteger("ID") & """>" & GetIssuePublishDate(cp, cs.GetInteger("ID")) & " " & cs.GetText("Name") & "</a>"
                            Stream &= "<br>"
                            If EncodeCopyNeeded Then
                                Stream &= cp.Utils.EncodeContentForWeb(cs.GetText("Overview"))
                            Else
                                Stream &= cs.GetText("Overview")
                            End If
                            'stream &=  "</TR>"
                            Call cs.GoNext()
                            If Colors = "#ffffff" Then
                                Colors = "#E0E0E0"
                            Else
                                Colors = "#ffffff"
                            End If
                        Loop
                    Else
                        BlockSearchForm = True
                        'stream &=  "<TR>"
                        Stream &= "<span class=""ccError"">" & cp.Site.GetText(SitePropertyNoNewsletterArchives, "There are currently no archived issues.") & "</span>"
                        'stream &=  "<TD><span class=""ccError"">" & cp.site.getText(SitePropertyNoNewsletterArchives, "There are currently no archived issues.", True) & "</span></TD>"
                        'stream &=  "</TR>"
                    End If
                    Call cs.Close()
                    'stream &=  "</TABLE>"
                End If
            End If
            If ButtonValue = FormButtonViewArchives Then
                '
                ' List search results of archive issues
                '
                'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                ThisSQL2 = " select NL.id, nl.name, nl.publishdate, nlp.AllowReadMore, nlp.Overview, nlp.Body, nlp.id as ThisID ,nlp.newsletterid, nlp.name as nlpname"
                ThisSQL2 = ThisSQL2 & " from newsletterissues nl, newsletterissuepages nlp"
                ThisSQL2 = ThisSQL2 & " Where (NL.ID = nlp.newsletterid)"
                If MonthSelected <> 0 Then
                    ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & MonthSelected
                End If
                If YearSelected <> 0 Then
                    ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & YearSelected
                End If
                If SearchKeywords <> "" Then
                    ThisSQL2 = ThisSQL2 & " and ((nlp.Body like '%" & SearchKeywords & "%' )or (nlp.name  like '%" & SearchKeywords & "%') or (nlp.Overview  like '%" & SearchKeywords & "%'))"
                End If
                ThisSQL2 = ThisSQL2 & "  ORDER BY PublishDate DESC"
                '
                Call cs.OpenSQL(ThisSQL2, RecordsPerPage, PageNumber)
                If Not cs.OK Then
                    Stream &= cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found")
                Else
                    Stream &= cp.Content.GetCopy("Newsletter Search Results Found", "Search results")
                    Do While cs.OK And RowCount < RecordsPerPage
                        SearchResult = cs.GetText("nlpname")
                        Dim thisid As Integer
                        thisid = cs.GetInteger("ID")
                        If Colors = "#E0E0E0" Then
                            Colors = "#ffffff"
                        Else
                            Colors = "#E0E0E0"
                        End If
                        'stream &=  "<tr><td  BGCOLOR= """ & Colors & """ style=""border-top:1px solid #c0c0c0;padding:20px;"">"
                        Stream &= "<div  class=""NewsletterBody"">"
                        Stream &= "<div  class=""Headline"">" & SearchResult & "</div>"
                        BriefCopyFileName = cs.GetText("Overview")
                        If BriefCopyFileName = "" Then
                            If cs.GetBoolean("AllowReadMore") Then
                                BriefCopyFileName = cs.GetText("Body")
                            Else
                                BriefCopyFileName = cp.Content.GetCopy("Newsletter Article Access Denied", "You do not have access to this article")
                            End If
                        End If
                        Stream &= "<div  class=""Overview"">" & cs.GetText("Overview") & "</div>"
                        'stream &=  cs.gettext( "Overview")
                        'stream &=  "<br><br>"
                        'stream &=  "</tr></td>"
                        If (cs.GetBoolean("AllowReadMore")) And (cs.GetText("Body") <> "") Then
                            qs = cp.Doc.RefreshQueryString
                            qs = cp.Utils.ModifyQueryString(qs, "formid", "400")
                            qs = cp.Utils.ModifyQueryString(qs, RequestNameIssuePageID, cs.GetInteger("ThisID"))
                            Stream &= "<a href=""?" & qs & """>"
                            Stream &= "Read More"
                            Stream &= "</a>"
                        End If
                        Stream &= "</div>"
                        Call cs.GoNext()
                        RowCount = RowCount + 1
                    Loop
                End If
                '
                If FileCount <> 0 Then
                    'stream &=  "<tr><td align=center>"
                    Stream &= "<div  class=""NewsletterBody""><div class=""GoToPageLine"">Go to Page&nbsp;&nbsp;"
                    Do While PageCount <= NumberofPages
                        'stream &=  "<a href=""" & Main.ServerPage & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """> Page " & (PageCount) & "</a>"
                        ' 1/1/09 - JK - alays linked to root path
                        Stream &= "<a href=""" & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """>" & (PageCount) & "</a>"
                        'stream &=  "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """>" & (PageCount) & "</a>"
                        PageCount = PageCount + 1
                        Stream &= "&nbsp;&nbsp;&nbsp;"
                    Loop
                    Stream &= "</div></div>"
                End If
                'stream &=  "</TABLE>"
            End If
            '
            If Not BlockSearchForm Then
                '
                ' Display search form
                '
                'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                'stream &=  "<tr>"
                'stream &=  "<td>"
                Stream &= cp.Content.GetCopy("Newsletter Search Copy", "<h2>Archive Search</h2>")
                'stream &=  "</td>"
                'stream &=  "</tr>"
                '
                'stream &=  "<tr>"
                'stream &=  "<td>"
                ' 1 ** drop down select list 2007 Issues (all issues in 2007)
                Stream &= "<div>" & cp.Html.SelectContent(RequestNameIssueID, "", ContentNameNewsletterIssues, "(Publishdate<" & cp.Db.EncodeSQLDate(Now) & ")AND(NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")") & " " & cp.Html.Button(FormButtonViewNewsLetter) & "</div>"
                ' ** need a button to view the newsletter
                'stream &=  "</td>"
                'stream &=  "</tr>"
                '
                Stream &= "<div>&nbsp;</div>"
                'stream &=  "<tr>"
                'stream &=  "<td>"
                Stream &= "<div>keyword search<br>"
                Stream &= cp.Html.InputText(RequestNameSearchKeywords, , , 50) & "</div>"
                'stream &=  "</td>"
                'stream &=  "</tr>"
                '
                'stream &=  "<tr>"
                'stream &=  "<td>"
                MonthString = ""
                MonthString &= "Month <select size=""1"" name=""" & RequestNameMonthSelectd & """>"
                MonthString &= "<option selected>Month</option>"
                For MonthCounter = 1 To 12
                    MonthString &= "<option "
                    MonthString &= "value=""" & MonthCounter & """>" & MonthName(MonthCounter) & "</option>"
                Next
                MonthString &= "</select> "
                '
                YearString = ""
                YearString &= "Year <select size=""1"" name=""" & RequestNameYearSelected & """>"
                YearString &= "<option selected>Year</option>"
                'For YearCounter = (Year(Now) - 5) To (Year(Now))
                For YearCounter = (Year(Now) - YearsWanted) To (Year(Now))
                    YearString &= "<option "
                    YearString &= "value=""" & YearCounter & """>" & YearCounter & "</option>"
                Next
                YearString &= "</select>"
                Stream &= "<div>&nbsp;</div>"
                Stream &= "<div>" & MonthString & "&nbsp;&nbsp;&nbsp;" & YearString & "&nbsp;&nbsp;&nbsp;&nbsp;" & cp.Html.Button(FormButtonViewArchives) & "</div>"
                'stream &=  "</td>"
                'stream &=  "</tr>"
                'stream &=  "</TABLE>"
            End If
            '
            Stream &= cp.Html.Hidden(RequestNameFormID, FormArchive)
            Stream &= cp.Html.Form(Stream)
            '
            GetArchiveList = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetArchiveList")
        End Function
        '
        Private Function GetFormRow(Innards As String) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            stream &= "<TR>"
            stream &= "<TD colspan=2 width=""60%"">" & Innards & "</TD>"
            stream &= "</TR>"
            '
            GetFormRow = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("DonationClass", "GetFormRow2")
        End Function
        '
        Private Function GetSpacer(Optional Height As Integer = 1, Optional Width As Integer = 1) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            Stream = "<img src=""/ccLib/images/spacer.gif"" width=""" & Width & """ height=""" & Height & """>"
            '
            GetSpacer = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("LeftSideNavigation", "GetSpacer")
        End Function
        '
        Private Function GetArticleAccess(cp As CPBaseClass, ArticleID As Integer, Optional GivenGroupID As Integer = 0) As Boolean
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim AccessFlag As Boolean
            Dim ThisTest As String
            '
            If GivenGroupID <> 0 Then
                Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                If Not cs.OK() Then
                    GetArticleAccess = True
                Else
                    Do While cs.OK()
                        If cs.GetInteger("GroupID") = GivenGroupID Then
                            GetArticleAccess = True
                        End If
                        Call cs.GoNext()
                    Loop
                End If
                Call cs.Close()
            Else
                If Not isManager Then
                    Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                    If Not cs.OK() Then
                        GetArticleAccess = True
                    Else
                        Do While cs.OK()
                            ThisTest = cs.GetText("GroupID")
                            '
                            '
                            If ThisTest <> "" Then
                                If cp.User.IsInGroup(ThisTest) Then
                                    GetArticleAccess = True
                                End If
                            End If
                            Call cs.GoNext()
                        Loop
                    End If
                    Call cs.Close()
                Else
                    GetArticleAccess = True
                End If
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetArticleAccess")
        End Function
        '
        Private Function GetIssuePublishDate(cp As CPBaseClass, IssueID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim IssueDate As String
            Dim Stream As String = ""
            '
            cs.Open(ContentNameNewsletterIssues, "ID=" & IssueID, , , "PublishDate")
            If cs.OK Then
                IssueDate = cs.GetDate("PublishDate")
                If IsDate(IssueDate) Then
                    Stream = MonthName(Month(IssueDate), True) & " " & Day(IssueDate) & ", " & Year(IssueDate)
                End If
            End If
            Call cs.Close()
            '
            '
            GetIssuePublishDate = Stream
            '
        End Function
        '
        Private Function GetEmailBody(cp As CPBaseClass, TemplateCopy As String, LocalGroupID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            Dim TemplateArray() As String
            Dim TemplateArrayCount As Integer
            Dim TemplateArrayPointer As Integer
            '
            Dim InnerTemplateArray() As String
            Dim InnerTemplateArrayCount As Integer
            Dim InnerTemplateArrayPointer As Integer
            '
            Dim InnerValue As String
            '
            Dim Navigation As New newsletterNavClass
            'Dim Mast As New MastClass
            '
            TemplateArray = Split(TemplateCopy, StringReplaceStart)
            TemplateArrayCount = UBound(TemplateArray) + 1
            For TemplateArrayPointer = 0 To TemplateArrayCount - 1
                If InStr(TemplateArray(TemplateArrayPointer), StringReplaceEnd) Then
                    InnerTemplateArray = Split(TemplateArray(TemplateArrayPointer), StringReplaceEnd)
                    InnerTemplateArrayCount = UBound(InnerTemplateArray)
                    InnerValue = InnerTemplateArray(0)
                    Select Case InnerValue
                        Case TemplateReplacementBody
                            Stream &= Replace(InnerValue, TemplateReplacementBody, GetNewsletterBodyOverview(cp, IssueID, IssuePageID, LocalGroupID))
                        Case TemplateReplacementNav
                            Stream &= Replace(InnerValue, TemplateReplacementNav, Navigation.GetContent(cp, "NavigationLayout=Vertical", LocalGroupID))
                    End Select
                    stream &= InnerTemplateArray(1)
                Else
                    stream &= TemplateArray(TemplateArrayPointer)
                End If
            Next
            '
            GetEmailBody = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetEmailBody")
        End Function
        '
        Private Function GetOverview(cp As CPBaseClass, PageID As Integer) As String
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            '
            Call cs.Open(ContentNameNewsletterIssuePages, "ID=" & cp.Db.EncodeSQLNumber(PageID), , , , , "Overview")
            If cs.OK() Then
                GetOverview = cs.GetText("Overview")
            End If
            Call cs.Close()
        End Function
        '
        Friend Function GetNewsletterBodyOverview(cp As CPBaseClass, IssueID As Integer, IssuePageID As Integer, Optional GivenGroupID As Integer = 0) As String
            'On Error GoTo ErrorTrap
            '
            Dim AddLink As String
            Dim Controls As String
            Dim IssueSQL As String
            Dim NewIssueId As Integer
            Dim MaxIssueID As Integer
            Dim Stream As String
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim Criteria As String
            Dim Link As String
            Dim HasArticleAccess As Boolean
            Dim SQL As String
            Dim TableList As String
            Dim CSTopics As Integer
            Dim MainSQL As String
            Dim PreviousCategoryName As String
            Dim CategoryName As String
            Dim RecordCount As Integer
            Dim cn As New newsletterCommonClass
            Dim AccessString As String
            Dim StoryID As Integer
            Dim StoryAccessString As String
            Dim Caption As String
            Dim FetchFlag As Boolean
            '
            Dim CategoryID As Integer
            Dim CS2 As CPCSBaseClass = cp.CSNew()
            '
            TableList = "NewsletterIssuePages "
            '
            Call cs.OpenRecord("Newsletter Issues", IssueID)
            If cs.OK() Then
                Stream &= cs.GetEditLink() & cs.getText("Cover")
            End If
            Call cs.Close()
            '
            If IssuePageID <> 0 Then
                Criteria = ""
                MainSQL = "" _
                    & " select p.categoryId,c.name as CategoryName" _
                    & " from NewsletterIssuePages p" _
                    & " left join NewsletterIssueCategories c on c.id=p.categoryId" _
                    & " where (p.ID=" & cp.Db.EncodeSQLNumber(IssuePageID) & ")" _
                    & ""
                'call cs.open(ContentNameNewsletterIssuePages, Criteria, "SortOrder,DateAdded")
            Else
                '
                FetchFlag = True
                '
                MainSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder, NIC.Name AS CategoryName"
                MainSQL = MainSQL & " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR"
                MainSQL = MainSQL & " Where (NIC.ID = NIR.CategoryID)"
                MainSQL = MainSQL & " AND (NIR.NewsletterIssueID=" & IssueID & ")"
                MainSQL = MainSQL & " AND (NIC.Active<>0)"
                MainSQL = MainSQL & " AND (NIR.Active<>0)"
                MainSQL = MainSQL & " ORDER BY NIR.SortOrder"
                '
                'Call cp.Site.TestPoint("MainSQL: " & MainSQL)
                'Call cs.OpenSQL(  MainSQL)
                '
            End If
            Call cp.Site.TestPoint("MainSQL: " & MainSQL)
            Call cs.OpenSQL(MainSQL)
            '
            Stream &= vbCrLf & "<!-- Start NewsletterBody -->" & vbCrLf
            Stream &= "<div class=""NewsletterBody"">"
            '
            If cs.OK() Then
                Do While cs.OK()
                    CategoryID = cs.GetInteger("CategoryID")
                    CategoryName = cs.GetText("CategoryName")
                    '
                    Call CS2.Open(ContentNameNewsletterIssuePages, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & IssueID & ")", "SortOrder")
                    If CS2.OK Then
                        '
                        ' there are stories under this topic, wrap in div to allow a story indent
                        '
                        Stream &= vbCrLf & "<div class=""NewsletterTopic"">"
                        If RecordCount <> 0 Then
                            If cp.User.IsAuthoring(ContentNameNewsletterIssuePages) Then
                                Stream &= cn.GetAdminHintWrapper(cp, "<a href=""" & WorkingQueryStringPlus & RequestNameIssueID & "=" & IssueID & "&" & RequestNameSortUp & "=" & CategoryID & """>[Move Up]</a> ")
                            End If
                        End If
                        Stream &= CategoryName
                        Stream &= "</div>"
                        '
                        Stream &= vbCrLf & "<div class=""NewsletterTopicStory"">"
                        Do While CS2.OK
                            Stream &= GetStoryOverview(CS2)
                            Call CS2.GoNext()
                        Loop
                        Stream &= "</div>"
                    End If
                    Call CS2.Close()
                    Call cs.GoNext()
                    RecordCount = RecordCount + 1
                Loop
            End If
            '
            Call cs.Close()
            '
            Stream &= GetUnrelatedStories(IssuePageID)
            '
            IssueSQL = " Select max(id) as MaxIssueID from newsletterissues"
            Call cs.OpenSQL(IssueSQL)
            If cs.OK Then
                MaxIssueID = cs.GetInteger("maxissueid")
            End If
            NewIssueId = MaxIssueID + 1
            Call cs.Close()
            '
            Stream &= "</div>"
            '
            Stream &= cp.Content.GetAddLink(ContentNameNewsletterIssuePages, "Newsletterid=" & IssueID, False, cp.User.IsEditingAnything)
            '
            Stream &= vbCrLf & "<!-- End NewsletterBody -->" & vbCrLf
            '
            GetNewsletterBodyOverview = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetNewsletterBodyOverview")
        End Function
        '

        Private Function GetUnrelatedStories(IssuePageID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim Criteria As String
            Dim cs As cpcsBaseClass = cp.csNew()
            Dim Caption As String
            Dim Stream As String
            '
            If IssuePageID = 0 Then
                Criteria = "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")"
                Call cs.open(ContentNameNewsletterIssuePages, Criteria, "SortOrder,DateAdded")
                If cs.ok() Then
                    Caption = cp.site.getText("Newsletter Caption Other Stories", "")
                    If Caption <> "" Then
                        stream &= vbCrLf & "<div class=""NewsletterTopic"">" & Caption & "</div>"
                    End If
                    Do While cs.ok()
                        stream &= GetStoryOverview(CS)
                        Call cs.gonext()
                    Loop
                End If
                Call cs.close()
            End If
            '
            GetUnrelatedStories = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetUnrelatedStories")
        End Function
        '
        Private Function GetStoryOverview(CS As CPCSBaseClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim StoryID As Integer
            Dim StoryAccessString As String
            Dim Stream As String
            Dim cn As New newsletterCommonClass
            Dim storyBookmark As String
            '
            StoryID = CS.getInteger("ID")
            storyBookmark = "story" & StoryID
            StoryAccessString = cn.GetArticleAccessString(cp, StoryID)
            '
            If StoryAccessString <> "" Then
                Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & StoryAccessString & """>"
            End If
            '
            Stream &= vbCrLf & "<div class=""Headline"" id=""" & storyBookmark & """>"
            If FormID <> FormEmail Then
                Stream &= CS.GetEditLink()
            End If
            Stream &= CS.getText("Name") & "</div>"
            'stream &=  "<a name=""" & storyBookmark & """>" &cs.getText( "Name") & "&nbsp;" & "</a></div>"
            If IssuePageID <> 0 Then
                If EncodeCopyNeeded Then
                    Stream &= "<div class=""Copy"">" & cp.Utils.EncodeContentForWeb(CS.GetText("Body ")) & "</div>"
                Else
                    Stream &= "<div class=""Copy"">" & CS.GetText("Body ") & "</div>"
                End If
            Else
                Stream &= "<div class=""Overview"">"
                Stream &= CS.getText("Overview")
                If CS.getBoolean("AllowReadMore") Then
                    ' 1/1/09 JK - always linked to root path
                    Stream &= "<div class=""ReadMore""><a href=""" & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & CS.getInteger("ID") & "&" & RequestNameFormID & "=" & FormDetails & """>Read More</a></div>"
                    'stream &=  "<div class=""ReadMore""><a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & cs.getInteger("ID") & "&" & RequestNameFormID & "=" & FormDetails & """>Read More</a></div>"
                    'Else
                    '    stream &=  "<div class=""ReadMore"">&nbsp;</div>"
                End If
                Stream &= "</div>"
            End If
            If StoryAccessString <> "" Then
                Stream &= "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
            End If
            '
            GetStoryOverview = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetStoryOverview")
        End Function
        '
        Private Function GetNewsletterBodyDetails(cp As CPBaseClass, cn As newsletterCommonClass, IssuePageID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim CSIssue As CPCSBaseClass = cp.CSNew()
            Dim rssChange As Boolean
            Dim expirationDate As Date
            Dim PublishDate As Date
            Dim Pos As Integer
            Dim recordDate As Date
            Dim Copy As String
            Dim Stream As String
            Dim Link As String
            Dim PrinterIcon As String
            Dim EmailIcon As String
            Dim storyName As String
            Dim storyOverview As String
            '
            PrinterIcon = "<img border=0 src=/ccLib/images/IconPrint.gif>"
            EmailIcon = "<img border=0 src=/ccLib/images/IconEmail.gif>"
            '
            If IssuePageID = 0 Then
                Stream = "<span class=""ccError"">The requested story is currently unavailable.</span>"
            Else
                Call cs.Open(ContentNameNewsletterIssuePages, "ID=" & IssuePageID)
                If cs.OK() Then
                    storyName = cs.GetText("name")
                    storyOverview = cs.GetText("Overview")
                    IssueID = cs.GetInteger("newsletterId")
                    '
                    Stream &= cs.GetEditLink()
                    If cs.GetBoolean("AllowPrinterPage") Then
                        Link = WorkingQueryStringPlus & RequestNameIssuePageID & "=" & IssuePageID & "&" & RequestNameFormID & "=" & FormDetails & "&ccIPage=l6d09a10sP"
                        Stream &= "<div class=""PrintIcon""><a target=_blank href=""" & Link & """>" & PrinterIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Printer Version</nobr></a></div>"
                    End If
                    If cs.GetBoolean("AllowEmailPage") Then
                        Link = "mailto:?SUBJECT=" & cp.Site.GetText("Email link subject", "A link to the " & cp.Site.DomainPrimary & " newsletter") & "&amp;BODY=http://" & cp.Site.DomainPrimary & cp.Site.AppRootPath & cp.Request.Page & Replace(WorkingQueryStringPlus, "&", "%26") & RequestNameIssuePageID & "=" & IssuePageID & "%26" & RequestNameFormID & "=" & FormDetails
                        Stream &= "<div class=""EmailIcon""><a target=_blank href=""" & Link & """>" & EmailIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Email this page</nobr></a></div>"
                    End If
                    Stream &= "<div class=""NewsletterBody"">"
                    Stream &= "<div class=""Headline"">" & cs.GetText("Name") & "</div>"
                    Stream &= "<div class=""Copy"">" & cs.GetText("Body") & "</div>"
                    Stream &= "</div>"
                    '
                    ' update RSS fields if empty
                    '
                    rssChange = False
                    If (IssueID <> 0) Then
                        If (cn.encodeMinDate(cs.GetDate("RSSDatePublish")) = Date.MinValue) Then
                            CSIssue.Open(ContentNameNewsletterIssues, "id=" & cp.Db.EncodeSQLNumber(IssueID))
                            If CSIssue.OK() Then
                                PublishDate = CSIssue.GetDate("publishDate")
                            End If
                            Call CSIssue.Close()
                            If (cn.encodeMinDate(PublishDate) <> Date.MinValue) Then
                                rssChange = True
                                Call cs.SetField("RSSDatePublish", PublishDate)
                            End If
                        End If
                    End If
                    '
                    If (storyName <> "") And (cs.GetText("RSSTitle") = "") Then
                        rssChange = True
                        Call cs.SetField("RSSTitle", storyName)
                    End If
                    '
                    If (storyOverview <> "") And (cs.GetText("RSSDescription") = "") Then
                        rssChange = True
                        Copy = cp.Utils.ConvertHTML2Text(storyOverview)
                        Call cs.SetField("RSSDescription", Copy)
                    End If
                    '
                    If (cs.GetText("RSSLink") = "") Then
                        Link = cp.Request.Link
                        If InStr(1, Link, cp.Site.GetText("adminUrl"), vbTextCompare) = 0 Then
                            Pos = InStr(1, Link, "?")
                            If Pos > 0 Then
                                Link = Left(Link, Pos - 1)
                            End If
                            Copy = cp.Doc.RefreshQueryString()
                            Copy = cp.Utils.ModifyQueryString(Copy, RequestNameIssuePageID, CStr(IssuePageID))
                            Copy = cp.Utils.ModifyQueryString(Copy, RequestNameFormID, FormDetails)
                            Copy = cp.Utils.ModifyQueryString(Copy, "method", "")
                            rssChange = True
                            Call cs.SetField("RSSLink", Link & "?" & Copy)
                        End If
                    End If
                    If rssChange Then
                        Call cp.Utils.ExecuteAddonAsProcess("RSS Feed Process")
                    End If
                End If
                Call cs.Close()
            End If
            '
            GetNewsletterBodyDetails = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetNewsletterBodyDetails")
        End Function
    End Class
End Namespace
