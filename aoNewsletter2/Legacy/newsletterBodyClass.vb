
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    Public Class newsletterBodyClass
        '
        Friend Function GetArchiveItemList(ByVal cp As CPBaseClass, ByVal cn As newsletterCommonClass, ByVal ButtonValue As String, ByVal currentIssueId As Integer, ByVal refreshQueryString As String, ByVal newsArchiveListItemLayout As String, ByVal NewsletterID As Integer) As String
            '
            Dim layout As New blockClass
            Dim recordTop As Integer
            Dim RecordsPerPage As Integer
            Dim archiveIssuesToDisplay As Integer
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim monthSelected As Integer
            Dim yearSelected As Integer
            Dim SearchKeywords As String
            '
            Dim link As String = ""
            Dim Stream As String = ""
            Dim Colors As String = ""
            Dim ThisSQL As String
            Dim ThisSQL2 As String = ""
            Dim MonthString As String
            Dim YearString As String
            Dim MonthCounter As Integer
            Dim YearCounter As Integer
            Dim storyName As String
            Dim NumberofPages As Integer
            Dim PageCount As Integer
            Dim sql2 As String
            Dim FileCount As Integer
            Dim storyOverview As String
            Dim RowCount As Integer
            Dim YearsWanted As Integer
            Dim BlockSearchForm As Boolean
            Dim qs As String
            Dim PageNumber As Integer
            Dim issueDate As Date
            Dim issueDateFormatted As String = ""
            Dim GoToPage As String = ""
            Dim storyBody As String = ""
            '
            BlockSearchForm = cp.Doc.GetBoolean("Block Archive Search Form")
            archiveIssuesToDisplay = cp.Doc.GetInteger("Archive Issues To Display")
            monthSelected = cp.Doc.GetInteger(RequestNameMonthSelectd)
            yearSelected = cp.Doc.GetInteger(RequestNameYearSelected)
            SearchKeywords = cp.Doc.GetText(RequestNameSearchKeywords)
            RecordsPerPage = cp.Site.GetInteger("Newsletter Search Results Records Per Page", "3")
            recordTop = cp.Doc.GetInteger(RequestNameRecordTop)
            '
            PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
            If PageNumber = 0 Then
                PageNumber = 1
            End If
            '
            YearsWanted = cp.Utils.EncodeInteger(cp.Site.GetText("Newsletter years wanted", "1"))
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
            sql2 = " select count(story.id) as count"
            sql2 = sql2 & " from newsletterissues nl, newsletterissuepages story"
            sql2 = sql2 & " Where (NL.ID = story.newsletterid)"
            sql2 = sql2 & " AND (NL.NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")"
            If monthSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & monthSelected
            End If
            If yearSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & yearSelected
            End If
            If SearchKeywords <> "" Then
                sql2 = sql2 & " and ((story.Body like '%" & SearchKeywords & "%' )or (story.name  like '%" & SearchKeywords & "%') or (story.Overview  like '%" & SearchKeywords & "%'))"
            End If
            If cs.OpenSQL(sql2) Then
                FileCount = cs.GetInteger("count")
                NumberofPages = CInt(FileCount / RecordsPerPage)
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
                If (monthSelected = 0) And (yearSelected = 0) Then
                    'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                    '
                    'ThisSQL = " SELECT  TOP 6 * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & cp.db.encodesqlNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    ThisSQL = " SELECT  TOP " & archiveIssuesToDisplay & " * " _
                        & " From NewsletterIssues " _
                        & " WHERE (PublishDate < { fn NOW() }) AND (ID <> " & currentIssueId & ") AND (NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ") " _
                        & " ORDER BY PublishDate DESC"
                    '
                    Call cs.OpenSQL(ThisSQL)
                    If cs.OK Then
                        Do While cs.OK
                            Call layout.load(newsArchiveListItemLayout)
                            issueDate = cs.GetDate("PublishDate")
                            If IsDate(issueDate) Then
                                issueDateFormatted = MonthName(Month(issueDate), True) & " " & Day(issueDate) & ", " & Year(issueDate)
                            End If
                            link = refreshQueryString
                            link = cp.Utils.ModifyQueryString(link, RequestNameIssueID, cs.GetInteger("ID").ToString())
                            Call layout.setClassInner("newsArchiveListCaption", cs.GetText("Name"))
                            Call layout.setClassInner("newsArchiveListOverview", cp.Utils.EncodeContentForWeb(cs.GetText("Overview")))
                            'Stream &= layout.GetHtml().Replace("?", "?" & link)
                            Stream &= Replace(layout.getHtml(), "href=""?""", "href=""?" & link & """")
                            Call cs.GoNext()
                        Loop
                    Else
                        BlockSearchForm = True
                        Call layout.load(newsArchiveListItemLayout)
                        Call layout.setClassInner("newsArchiveListCaption", "<span class=""ccError"">" & cp.Site.GetText(SitePropertyNoNewsletterArchives, "There are currently no archived issues.") & "</span>")
                        Call layout.setClassInner("newsArchiveListOverview", "")
                        Stream &= layout.getHtml()
                    End If
                    Call cs.Close()
                End If
            End If
            If ButtonValue = FormButtonViewArchives Then
                '
                ' List search results of archive issues
                '
                cp.Utils.AppendLog("test.log", cp.Doc.GetInteger("newsletter").ToString())

                'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                ThisSQL2 = " select NL.id, nl.name, nl.publishdate, story.AllowReadMore, story.Overview, story.Body, story.id as ThisID ,story.newsletterid, story.name as storyName"
                ThisSQL2 = ThisSQL2 & " from newsletterissues nl, newsletterissuepages story"
                ThisSQL2 = ThisSQL2 & " Where (NL.ID = story.newsletterid) "
                ThisSQL2 = ThisSQL2 & " and nl.NewsletterID=" & NewsletterID & " " ' 01/13/2017 Search only in the same NewsletterID
                If monthSelected <> 0 Then
                    ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & monthSelected
                End If
                If yearSelected <> 0 Then
                    ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & yearSelected
                End If
                If SearchKeywords <> "" Then
                    ThisSQL2 = ThisSQL2 & " and ((story.Body like '%" & SearchKeywords & "%' )or (story.name  like '%" & SearchKeywords & "%') or (story.Overview  like '%" & SearchKeywords & "%'))"
                End If
                ThisSQL2 = ThisSQL2 & "  ORDER BY PublishDate DESC"
                '
                'Call cs.OpenSQL(ThisSQL2, "", RecordsPerPage, PageNumber)
                Call cs.OpenSQL(ThisSQL2, "")
                If Not cs.OK Then
                    Call layout.load(newsArchiveListItemLayout)
                    Call layout.setClassInner("newsArchiveListCaption", "No results were found")
                    'Call layout.SetClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found"))
                    Call layout.setClassInner("newsArchiveListOverview", "")
                    Stream &= Replace(layout.getHtml(), "?", "?" & cp.Utils.ModifyQueryString(refreshQueryString, RequestNameFormID, FormArchive.ToString(), True))  'layout.GetHtml()
                    'Stream &= cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found")
                Else
                    Call layout.load(newsArchiveListItemLayout)
                    Call layout.setClassInner("newsArchiveListCaption", "Search results")
                    'Call layout.SetClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search Results Found", "Search results"))
                    Call layout.setClassInner("newsArchiveListOverview", "")
                    Stream &= Replace(layout.getHtml(), "?", "?" & cp.Utils.ModifyQueryString(refreshQueryString, RequestNameFormID, FormArchive.ToString(), True)) 'layout.GetHtml()
                    Do While cs.OK And RowCount < RecordsPerPage
                        storyName = cs.GetText("storyName")
                        storyOverview = cs.GetText("Overview")
                        storyBody = cs.GetText("body")
                        If storyOverview = "" Then
                            If Not cn.isBlank(cp, storyBody) Then
                                'if cs.GetBoolean("AllowReadMore") Then
                                storyOverview = storyBody
                            Else
                                storyOverview = cp.Content.GetCopy("Newsletter Article Access Denied", "You do not have access to this article")
                            End If
                        End If
                        qs = refreshQueryString
                        ' 01/12/2017 Dwayne request change the link to the full history
                        'qs = cp.Utils.ModifyQueryString(qs, "formid", FormCover.ToString())
                        qs = cp.Utils.ModifyQueryString(qs, "formid", FormDetails.ToString())
                        qs = cp.Utils.ModifyQueryString(qs, RequestNameStoryId, cs.GetInteger("ThisID").ToString())
                        Call layout.load(newsArchiveListItemLayout)
                        Call layout.setClassInner("newsArchiveListCaption", storyName)
                        Call layout.setClassInner("newsArchiveListOverview", storyOverview)
                        If layout.getHtml().Contains("?") Then
                            ' cp.Utils.AppendLog("Test2.log", layout.getHtml())
                        End If
                        Stream &= Replace(layout.getHtml(), "href=""?""", "href=""?" & qs & """")
                        Call cs.GoNext()
                        RowCount = RowCount + 1
                    Loop
                End If
                '
                ' 01/13/2017 comment pagination
                '
                'If FileCount <> 0 Then
                '    GoToPage = ""
                '    Do While PageCount <= NumberofPages
                '        qs = refreshQueryString
                '        qs = cp.Utils.ModifyQueryString(qs, RequestNameButtonValue, FormButtonViewArchives)
                '        qs = cp.Utils.ModifyQueryString(qs, RequestNamePageNumber, PageCount.ToString())
                '        qs = cp.Utils.ModifyQueryString(qs, RequestNameSearchKeywords, SearchKeywords)
                '        GoToPage &= "<a href=""?" & qs & """>" & (PageCount) & "</a>"
                '        PageCount = PageCount + 1
                '        GoToPage &= "&nbsp;&nbsp;&nbsp;"
                '    Loop
                '    Call layout.Load(newsArchiveListItemLayout)
                '    Call layout.SetClassInner("newsArchiveListCaption", GoToPage)
                '    Call layout.SetClassInner("newsArchiveListOverview", "")
                '    Stream &= layout.GetHtml()
                'End If
            End If
            '
            If Not BlockSearchForm Then
                '
                ' Display search form
                '
                Dim searchForm As String = ""
                searchForm &= "<h2>Archive Search</h2>"
                'searchForm &= cp.Content.GetCopy("Newsletter Search Copy", "<h2>Archive Search</h2>")
                searchForm &= "<div>" & cp.Html.SelectContent(RequestNameIssueID, "", ContentNameNewsletterIssues, "(Publishdate<" & cp.Db.EncodeSQLDate(Now) & ")AND(NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")") & "</div>"
                'searchForm &= "<div>" & cp.Html.SelectContent(RequestNameIssueID, "", ContentNameNewsletterIssues, "(Publishdate<" & cp.Db.EncodeSQLDate(Now) & ")AND(NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")") & " " & cp.Html.Button(FormButtonViewNewsLetter) & "</div>"
                searchForm &= "<div>&nbsp;</div>"
                searchForm &= "<div>keyword search<br>"
                searchForm &= cp.Html.InputText(RequestNameSearchKeywords, , , "50") & "</div>"
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
                For YearCounter = (Year(Now) - YearsWanted) To (Year(Now))
                    YearString &= "<option "
                    YearString &= "value=""" & YearCounter & """>" & YearCounter & "</option>"
                Next
                YearString &= "</select>"
                searchForm &= "<div>&nbsp;</div>"
                searchForm &= "<div>" & MonthString & "&nbsp;&nbsp;&nbsp;" & YearString & "&nbsp;&nbsp;&nbsp;&nbsp;" & cp.Html.Button("button", FormButtonViewArchives) & "</div>"
                searchForm &= "<div>&nbsp;</div>"
                searchForm &= cp.Html.Hidden(RequestNameFormID, FormArchive.ToString())
                searchForm = cp.Html.Form(searchForm)
                '
                Call layout.load(newsArchiveListItemLayout)
                Call layout.setClassInner("newsArchiveListCaption", "")
                Call layout.setClassInner("newsArchiveListOverview", searchForm)
                Stream &= layout.getHtml()
            End If
            '
            '
            GetArchiveItemList = Stream
        End Function
        '
        '
        Friend Function GetSearchItemList(ByVal cp As CPBaseClass, ByVal cn As newsletterCommonClass, ByVal ButtonValue As String, ByVal issueId As Integer, ByVal refreshQueryString As String, ByVal newsArchiveListItemLayout As String) As String
            '
            Dim layout As New blockClass()
            Dim recordTop As Integer
            Dim RecordsPerPage As Integer
            Dim archiveIssuesToDisplay As Integer
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim NewsletterID As Integer
            Dim monthSelected As Integer
            Dim yearSelected As Integer
            Dim SearchKeywords As String
            '
            Dim link As String = ""
            Dim Stream As String = ""
            Dim Colors As String = ""
            Dim ThisSQL As String
            Dim ThisSQL2 As String = ""
            Dim MonthString As String
            Dim YearString As String
            Dim MonthCounter As Integer
            Dim YearCounter As Integer
            Dim storyName As String
            Dim NumberofPages As Integer
            Dim PageCount As Integer
            Dim sql2 As String
            Dim FileCount As Integer
            Dim storyOverview As String
            Dim RowCount As Integer
            Dim YearsWanted As Integer
            Dim BlockSearchForm As Boolean
            Dim qs As String
            Dim PageNumber As Integer
            Dim issueDate As Date
            Dim issueDateFormatted As String = ""
            Dim GoToPage As String = ""
            Dim storyBody As String = ""
            '
            archiveIssuesToDisplay = cp.Doc.GetInteger("Archive Issues To Display")
            monthSelected = cp.Doc.GetInteger(RequestNameMonthSelectd)
            yearSelected = cp.Doc.GetInteger(RequestNameYearSelected)
            SearchKeywords = cp.Doc.GetText(RequestNameSearchKeywords)
            RecordsPerPage = cp.Site.GetInteger("Newsletter Search Results Records Per Page", "3")
            recordTop = cp.Doc.GetInteger(RequestNameRecordTop)
            '
            PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
            If PageNumber = 0 Then
                PageNumber = 1
            End If
            '
            YearsWanted = cp.Utils.EncodeInteger(cp.Site.GetText("Newsletter years wanted", "1"))
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
            sql2 = " select count(story.id) as count"
            sql2 = sql2 & " from newsletterissues nl, newsletterissuepages story"
            sql2 = sql2 & " Where (NL.ID = story.newsletterid)"
            sql2 = sql2 & " AND (NL.NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")"
            If monthSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & monthSelected
            End If
            If yearSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & yearSelected
            End If
            If SearchKeywords <> "" Then
                sql2 = sql2 & " and ((story.Body like '%" & SearchKeywords & "%' )or (story.name  like '%" & SearchKeywords & "%') or (story.Overview  like '%" & SearchKeywords & "%'))"
            End If
            If cs.OpenSQL(sql2) Then
                FileCount = cs.GetInteger("count")
                NumberofPages = CInt(FileCount / RecordsPerPage)
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
                If (monthSelected = 0) And (yearSelected = 0) Then
                    'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                    '
                    'ThisSQL = " SELECT  TOP 6 * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & cp.db.encodesqlNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    ThisSQL = " SELECT  TOP " & archiveIssuesToDisplay & " * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & issueId & ") AND (NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    '
                    Call cs.OpenSQL(ThisSQL)
                    If cs.OK Then
                        Do While cs.OK
                            Call layout.load(newsArchiveListItemLayout)
                            issueDate = cs.GetDate("PublishDate")
                            If IsDate(issueDate) Then
                                issueDateFormatted = MonthName(Month(issueDate), True) & " " & Day(issueDate) & ", " & Year(issueDate)
                            End If
                            link = refreshQueryString
                            link = cp.Utils.ModifyQueryString(link, RequestNameIssueID, cs.GetInteger("ID").ToString())
                            Call layout.setClassInner("newsArchiveListCaption", cs.GetText("Name"))
                            Call layout.setClassInner("newsArchiveListOverview", cp.Utils.EncodeContentForWeb(cs.GetText("Overview")))
                            Stream &= Replace(layout.getHtml(), "?", "?" & link)
                            Call cs.GoNext()
                        Loop
                    Else
                        BlockSearchForm = True
                        Call layout.load(newsArchiveListItemLayout)
                        Call layout.setClassInner("newsArchiveListCaption", "<span class=""ccError"">" & cp.Site.GetText(SitePropertyNoNewsletterArchives, "There are currently no archived issues.") & "</span>")
                        Call layout.setClassInner("newsArchiveListOverview", "")
                        Stream &= layout.getHtml()
                    End If
                    Call cs.Close()
                End If
            End If
            If ButtonValue = FormButtonViewArchives Then
                '
                ' List search results of archive issues
                '
                'stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                ThisSQL2 = " select NL.id, nl.name, nl.publishdate, story.AllowReadMore, story.Overview, story.Body, story.id as ThisID ,story.newsletterid, story.name as storyName"
                ThisSQL2 = ThisSQL2 & " from newsletterissues nl, newsletterissuepages story"
                ThisSQL2 = ThisSQL2 & " Where (NL.ID = story.newsletterid)"
                If monthSelected <> 0 Then
                    ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & monthSelected
                End If
                If yearSelected <> 0 Then
                    ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & yearSelected
                End If
                If SearchKeywords <> "" Then
                    ThisSQL2 = ThisSQL2 & " and ((story.Body like '%" & SearchKeywords & "%' )or (story.name  like '%" & SearchKeywords & "%') or (story.Overview  like '%" & SearchKeywords & "%'))"
                End If
                ThisSQL2 = ThisSQL2 & "  ORDER BY PublishDate DESC"
                '
                Call cs.OpenSQL(ThisSQL2, "", RecordsPerPage, PageNumber)
                If Not cs.OK Then
                    Call layout.load(newsArchiveListItemLayout)
                    Call layout.setClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found"))
                    Call layout.setClassInner("newsArchiveListOverview", "")
                    Stream &= layout.getHtml()
                    'Stream &= cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found")
                Else
                    Call layout.load(newsArchiveListItemLayout)
                    Call layout.setClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search Results Found", "Search results"))
                    Call layout.setClassInner("newsArchiveListOverview", "")
                    Stream &= layout.getHtml()
                    Do While cs.OK And RowCount < RecordsPerPage
                        storyName = cs.GetText("storyName")
                        storyOverview = cs.GetText("Overview")
                        storyBody = cs.GetText("body")
                        If storyOverview = "" Then
                            If Not cn.isBlank(cp, storyBody) Then
                                ' if cs.GetBoolean("AllowReadMore") Then
                                storyOverview = storyBody
                            Else
                                storyOverview = cp.Content.GetCopy("Newsletter Article Access Denied", "You do not have access to this article")
                            End If
                        End If
                        qs = refreshQueryString
                        qs = cp.Utils.ModifyQueryString(qs, "formid", "400")
                        qs = cp.Utils.ModifyQueryString(qs, RequestNameStoryId, cs.GetInteger("ThisID").ToString())
                        Call layout.load(newsArchiveListItemLayout)
                        Call layout.setClassInner("newsArchiveListCaption", storyName)
                        Call layout.setClassInner("newsArchiveListOverview", storyOverview)
                        Stream &= Replace(layout.getHtml(), "?", "?" & qs)
                        Call cs.GoNext()
                        RowCount = RowCount + 1
                    Loop
                End If
                '
                If FileCount <> 0 Then
                    GoToPage = ""
                    Do While PageCount <= NumberofPages
                        qs = refreshQueryString
                        qs = cp.Utils.ModifyQueryString(qs, RequestNameButtonValue, FormButtonViewArchives)
                        qs = cp.Utils.ModifyQueryString(qs, RequestNamePageNumber, PageCount.ToString())
                        qs = cp.Utils.ModifyQueryString(qs, RequestNameSearchKeywords, SearchKeywords)
                        GoToPage &= "<a href=""?" & qs & """>" & (PageCount) & "</a>"
                        PageCount = PageCount + 1
                        GoToPage &= "&nbsp;&nbsp;&nbsp;"
                    Loop
                    Call layout.load(newsArchiveListItemLayout)
                    Call layout.setClassInner("newsArchiveListCaption", GoToPage)
                    Call layout.setClassInner("newsArchiveListOverview", "")
                    Stream &= layout.getHtml()
                End If
            End If
            '
            If Not BlockSearchForm Then
                '
                ' Display search form
                '
                Dim searchForm As String = ""
                searchForm &= cp.Content.GetCopy("Newsletter Search Copy", "<h2>Archive Search</h2>")
                searchForm &= "<div>" & cp.Html.SelectContent(RequestNameIssueID, "", ContentNameNewsletterIssues, "(Publishdate<" & cp.Db.EncodeSQLDate(Now) & ")AND(NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")") & " " & cp.Html.Button(FormButtonViewNewsLetter) & "</div>"
                searchForm &= "<div>&nbsp;</div>"
                searchForm &= "<div>keyword search<br>"
                searchForm &= cp.Html.InputText(RequestNameSearchKeywords, , , "50") & "</div>"
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
                For YearCounter = (Year(Now) - YearsWanted) To (Year(Now))
                    YearString &= "<option "
                    YearString &= "value=""" & YearCounter & """>" & YearCounter & "</option>"
                Next
                YearString &= "</select>"
                searchForm &= "<div>&nbsp;</div>"
                searchForm &= "<div>" & MonthString & "&nbsp;&nbsp;&nbsp;" & YearString & "&nbsp;&nbsp;&nbsp;&nbsp;" & cp.Html.Button(FormButtonViewArchives) & "</div>"
                searchForm &= cp.Html.Hidden(RequestNameFormID, FormArchive.ToString())
                searchForm &= cp.Html.Form(searchForm)
                '
                Call layout.load(newsArchiveListItemLayout)
                Call layout.setClassInner("newsArchiveListCaption", "")
                Call layout.setClassInner("newsArchiveListOverview", searchForm)
                Stream &= layout.getHtml()
            End If
            '
            '
            GetSearchItemList = Stream
        End Function
        '
        Private Function GetFormRow(ByVal Innards As String) As String
            Dim Stream As String = ""
            '
            Stream &= "<TR>"
            Stream &= "<TD colspan=2 width=""60%"">" & Innards & "</TD>"
            Stream &= "</TR>"
            GetFormRow = Stream
        End Function
        '
        Private Function GetSpacer(Optional ByVal Height As Integer = 1, Optional ByVal Width As Integer = 1) As String
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
        ''
        'Private Function GetArticleAccess(cp As CPBaseClass, ArticleID As Integer, isManager As Boolean, Optional GivenGroupID As Integer = 0) As Boolean
        '    'On Error GoTo ErrorTrap
        '    '
        '    Dim cs As CPCSBaseClass = cp.CSNew()
        '    Dim AccessFlag As Boolean
        '    Dim ThisTest As String
        '    '
        '    If GivenGroupID <> 0 Then
        '        Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
        '        If Not cs.OK() Then
        '            GetArticleAccess = True
        '        Else
        '            Do While cs.OK()
        '                If cs.GetInteger("GroupID") = GivenGroupID Then
        '                    GetArticleAccess = True
        '                End If
        '                Call cs.GoNext()
        '            Loop
        '        End If
        '        Call cs.Close()
        '    Else
        '        If Not isManager Then
        '            Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
        '            If Not cs.OK() Then
        '                GetArticleAccess = True
        '            Else
        '                Do While cs.OK()
        '                    ThisTest = cs.GetText("GroupID")
        '                    '
        '                    '
        '                    If ThisTest <> "" Then
        '                        If cp.User.IsInGroup(ThisTest) Then
        '                            GetArticleAccess = True
        '                        End If
        '                    End If
        '                    Call cs.GoNext()
        '                Loop
        '            End If
        '            Call cs.Close()
        '        Else
        '            GetArticleAccess = True
        '        End If
        '    End If
        '    '
        '    'Exit Function
        '    'ErrorTrap:
        '    'Call HandleError(cp, ex, "GetArticleAccess")
        'End Function
        '
        'Private Function GetIssuePublishDate(ByVal cp As CPBaseClass, ByVal IssueID As Integer) As String
        '    'On Error GoTo ErrorTrap
        '    '
        '    Dim cs As CPCSBaseClass = cp.CSNew()
        '    Dim IssueDate As String
        '    Dim Stream As String = ""
        '    '
        '    cs.Open(ContentNameNewsletterIssues, "ID=" & IssueID, , , "PublishDate")
        '    If cs.OK Then
        '        IssueDate = cs.GetDate("PublishDate")
        '        If IsDate(IssueDate) Then
        '            Stream = MonthName(Month(IssueDate), True) & " " & Day(IssueDate) & ", " & Year(IssueDate)
        '        End If
        '    End If
        '    Call cs.Close()
        '    '
        '    '
        '    GetIssuePublishDate = Stream
        '    '
        'End Function
        ''
        Friend Function GetCoverContent(ByVal cp As CPBaseClass, ByVal IssueID As Integer, ByVal storyId As Integer, ByVal refreshQueryString As String, ByVal formid As Integer, ByVal newsCoverStoryItem As String, ByVal newsCoverCategoryItem As String, ByVal isEditing As Boolean, ByRef return_Sponsor As String, ByRef return_publishDate As Date, ByRef return_tagLine As String) As String
            Dim returnHtmlItemList As String = ""
            Try
                '
                Dim layout As New blockClass
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim Criteria As String
                Dim TableList As String
                Dim MainSQL As String
                Dim CategoryName As String
                Dim RecordCount As Integer
                Dim cn As New newsletterCommonClass
                Dim FetchFlag As Boolean
                Dim CategoryID As Integer
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim qs As String
                Dim cover As String
                '
                TableList = "NewsletterIssuePages "
                '
                Call openRecord(cp, cs, "Newsletter Issues", IssueID)
                If cs.OK() Then
                    cover = cs.GetText("Cover")
                    return_Sponsor = cs.GetText("sponsor")
                    return_tagLine = cs.GetText("tagLine")
                    return_publishDate = genericController.encodeMinDate(cs.GetDate("publishDate"))
                    If cover.Length > 50 Then
                        returnHtmlItemList = GetCoverStoryItemLayout(cp, newsCoverStoryItem, "", "", "", cover, "", "", "", "")
                    End If
                End If
                Call cs.Close()
                '
                If storyId <> 0 Then
                    Criteria = ""
                    MainSQL = "" _
                        & " select p.categoryId,c.name as CategoryName" _
                        & " from NewsletterIssuePages p" _
                        & " left join NewsletterIssueCategories c on c.id=p.categoryId" _
                        & " where (p.ID=" & cp.Db.EncodeSQLNumber(storyId) & ")" _
                        & ""
                    'call cs.open(ContentNameNewsletterStories, Criteria, "SortOrder,DateAdded")
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
                If cs.OK() Then
                    Do While cs.OK()
                        CategoryID = cs.GetInteger("CategoryID")
                        '
                        Call CS2.Open(ContentNameNewsletterStories, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & IssueID & ")", "SortOrder,id")
                        If CS2.OK Then
                            '
                            ' there are stories under this topic, wrap in div to allow a story indent
                            '
                            Call layout.load(newsCoverCategoryItem)
                            CategoryName = cs.GetText("CategoryName")
                            If isEditing And (RecordCount <> 0) Then
                                qs = refreshQueryString
                                qs = cp.Utils.ModifyQueryString(qs, RequestNameIssueID, IssueID.ToString())
                                qs = cp.Utils.ModifyQueryString(qs, RequestNameSortUp, CategoryID.ToString())
                                CategoryName &= "&nbsp;<a href=""?" & qs & """><span style=""font-family:helvetica,arial,san-serif;font-weight:Normal;font-size:13px;text-decoration:none;"">[Move Up]</span></a> "
                            End If
                            Call layout.setClassInner("newsCoverCategoryItem", CategoryName)
                            returnHtmlItemList &= layout.getHtml()
                            '
                            Do While CS2.OK
                                returnHtmlItemList &= GetCoverStoryItem(cp, CS2, formid, refreshQueryString, newsCoverStoryItem, isEditing)
                                Call CS2.GoNext()
                            Loop
                            'returnHtmlItemList &= layout.GetHtml()
                        End If
                        Call CS2.Close()
                        Call cs.GoNext()
                        RecordCount = RecordCount + 1
                    Loop
                End If
                '
                Call cs.Close()
                '
                Criteria = "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")"
                Call cs.Open(ContentNameNewsletterStories, Criteria, "SortOrder,DateAdded")
                If cs.OK() Then
                    'Caption = cp.Site.GetText("Newsletter Caption Other Stories", "")
                    'If Caption <> "" Then
                    '    Stream &= vbCrLf & "<div class=""NewsletterTopic"">" & Caption & "</div>"
                    'End If
                    Do While cs.OK()
                        returnHtmlItemList &= GetCoverStoryItem(cp, cs, formid, refreshQueryString, newsCoverStoryItem, isEditing)
                        Call cs.GoNext()
                    Loop
                End If
                Call cs.Close()
                '
                If isEditing Then
                    Call layout.load(newsCoverStoryItem)
                    Call layout.setClassInner("newsCoverListCaption", cp.Content.GetAddLink(ContentNameNewsletterStories, "Newsletterid=" & IssueID, False, cp.User.IsEditingAnything) & "Add a story to this issue")
                    Call layout.setClassInner("newsCoverListOverview", "")
                    Call layout.setClassInner("newsCoverListReadMore", "")
                    Call layout.setClassInner("infographicBox", "")
                    returnHtmlItemList &= layout.getHtml()
                End If
            Catch ex As Exception
                Call handleError(cp, ex, "GetNewsletterBodyOverview")
            End Try
            Return returnHtmlItemList
        End Function
        ''

        'Private Function GetUnrelatedStories(ByVal cp As CPBaseClass, ByVal IssuePageID As Integer, ByVal IssueID As Integer, ByVal formId As Integer, ByVal refreshQueryString As String, ByVal newsCoverStoryItem As String) As String
        '    Dim returnHtml As String = ""
        '    Try

        '        'On Error GoTo ErrorTrap
        '        '
        '        Dim Criteria As String
        '        Dim cs As CPCSBaseClass = cp.CSNew()
        '        Dim Caption As String
        '        '
        '        If IssuePageID = 0 Then
        '            Criteria = "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")"
        '            Call cs.Open(ContentNameNewsletterStories, Criteria, "SortOrder,DateAdded")
        '            If cs.OK() Then
        '                'Caption = cp.Site.GetText("Newsletter Caption Other Stories", "")
        '                'If Caption <> "" Then
        '                '    Stream &= vbCrLf & "<div class=""NewsletterTopic"">" & Caption & "</div>"
        '                'End If
        '                Do While cs.OK()
        '                    returnHtml &= GetStoryOverview(cp, cs, formId, IssuePageID, refreshQueryString, newsCoverStoryItem)
        '                    Call cs.GoNext()
        '                Loop
        '            End If
        '            Call cs.Close()
        '        End If
        '    Catch ex As Exception
        '        Call handleError(cp, ex, "GetUnrelatedStories")
        '    End Try
        '    Return returnHtml
        'End Function
        '
        Private Function GetCoverStoryItem(ByVal cp As CPBaseClass, ByVal CSStories As CPCSBaseClass, ByVal formId As Integer, ByVal refreshQueryString As String, ByVal newsCoverStoryItem As String, isEditing As Boolean) As String
            Dim returnhtml As String = ""
            Try
                '
                Dim StoryID As Integer
                Dim StoryAccessString As String
                Dim cn As New newsletterCommonClass
                Dim storyBookmark As String
                Dim caption As String = ""
                Dim readMoreLink As String = ""
                Dim readMore As String = ""
                Dim overview As String = ""
                Dim storyBody As String = ""
                Dim coverInfographicthumbnail As String = ""
                Dim coverInfographic As String = ""
                Dim coverInfographicUrl As String = ""
                '
                StoryID = CSStories.GetInteger("ID")
                coverInfographicthumbnail = CSStories.GetText("coverInfographicthumbnail")
                coverInfographic = CSStories.GetText("coverInfographic")
                coverInfographicUrl = CSStories.GetText("coverInfographicUrl")
                storyBookmark = "story" & StoryID
                '
                StoryAccessString = cn.GetArticleAccessString(cp, StoryID)
                '
                If formId <> FormEmail Then
                    caption &= CSStories.GetEditLink()
                End If
                caption = "<span id=""" & storyBookmark & """>" & CSStories.GetText("Name") & "</span>"
                If isEditing Then
                    caption = CSStories.GetEditLink() & caption
                End If
                overview &= cp.Utils.EncodeContentForWeb(CSStories.GetText("Overview"))
                storyBody = CSStories.GetText("body")
                If Not cn.isBlank(cp, storyBody) Then
                    readMoreLink = refreshQueryString
                    readMoreLink = cp.Utils.ModifyQueryString(readMoreLink, RequestNameStoryId, StoryID.ToString())
                    readMoreLink = cp.Utils.ModifyQueryString(readMoreLink, RequestNameFormID, FormDetails.ToString())
                End If
                returnhtml = GetCoverStoryItemLayout(cp, newsCoverStoryItem, StoryAccessString, storyBookmark, caption, overview, readMoreLink, coverInfographicthumbnail, coverInfographic, coverInfographicUrl)
            Catch ex As Exception
                Call handleError(cp, ex, "getStoryOverview")
            End Try
            Return returnhtml
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' Populate an instance of the cover item template
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="newsCoverStoryItem"></param>
        ''' <param name="StoryAccessString"></param>
        ''' <param name="storyBookmark"></param>
        ''' <param name="caption"></param>
        ''' <param name="overview"></param>
        ''' <param name="readMoreLink"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetCoverStoryItemLayout(ByVal cp As CPBaseClass, newsCoverStoryItem As String, StoryAccessString As String, storyBookmark As String, caption As String, overview As String, readMoreLink As String, coverinfographicThumbnail As String, coverinfographic As String, coverInfographicUrl As String) As String
            Dim returnhtml As String = ""
            Try
                '
                Dim layout As New blockClass
                Dim cn As New newsletterCommonClass
                Dim readMore As String = ""
                Dim storyBody As String = ""
                Dim img As String = ""
                '
                Call layout.load(newsCoverStoryItem)
                '
                If String.IsNullOrEmpty(coverinfographicThumbnail) Then
                    '
                    ' no infographic
                    '
                    layout.setClassOuter("infographicBox", "")
                Else
                    img = "<img src=""" & cp.Site.FilePath & coverinfographicThumbnail & """ alt=""View the infographic"" class=""banner"" width=""100%"">"
                    If String.IsNullOrEmpty(coverinfographic) Then
                        '
                        ' no image
                        '
                        If String.IsNullOrEmpty(coverInfographicUrl) Then
                            layout.setClassInner("infographImage", img)
                        Else
                            If coverInfographicUrl.IndexOf("://") < 0 Then
                                coverInfographicUrl = "http://" & coverInfographicUrl
                            End If
                            layout.setClassInner("infographImage", "<a href=""" & coverInfographicUrl & """ target=""_blank"">" & img & "</a>")
                        End If
                    Else
                        '
                        ' linked thumbnail
                        '
                        layout.setClassInner("infographImage", "<a href=""" & cp.Site.FilePath & coverinfographic & """ target=""_blank"">" & img & "</a>")
                    End If
                End If
                If String.IsNullOrEmpty(coverinfographic) Then
                    layout.setClassOuter("infographLink", "")
                Else
                    layout.setClassInner("infographLink", "<a href=""" & cp.Site.FilePath & coverinfographic & """ target=""_blank"">View the infographic online.</a>")
                End If
                If StoryAccessString <> "" Then
                    Call layout.prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & StoryAccessString & """>")
                End If
                If String.IsNullOrEmpty(caption) Then
                    Call layout.setClassOuter("newsCoverListCaption", "")
                Else
                    Call layout.setClassInner("newsCoverListCaption", caption)
                End If
                If String.IsNullOrEmpty(overview) Then
                    Call layout.setClassOuter("newsCoverListOverview", "")
                Else
                    Call layout.setClassInner("newsCoverListOverview", overview)
                End If

                If (String.IsNullOrEmpty(readMoreLink)) Then
                    Call layout.setClassOuter("newsCoverListReadMore", "")
                Else
                    readMore = layout.getClassInner("newsCoverListReadMore")
                    readMore = readMore.Replace("?", "?" & readMoreLink)
                    readMore = readMore.Replace("#", "?" & readMoreLink)
                    Call layout.setClassInner("newsCoverListReadMore", readMore)
                End If
                If StoryAccessString <> "" Then
                    Call layout.append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                End If
                '
                returnhtml = layout.getHtml()
            Catch ex As Exception
                Call handleError(cp, ex, "getStoryOverview")
            End Try
            Return returnhtml
        End Function
        '
        Friend Function GetStory(ByVal cp As CPBaseClass, ByVal cn As newsletterCommonClass, ByVal storyId As Integer, ByVal IssueID As Integer, ByVal refreshQueryString As String, ByVal newsBody As String, isEditing As Boolean) As String
            Dim returnHtml As String = ""
            Try
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim CSIssue As CPCSBaseClass = cp.CSNew()
                Dim rssChange As Boolean
                Dim PublishDate As Date
                Dim Pos As Integer
                Dim Copy As String
                Dim Link As String
                Dim PrinterIcon As String
                Dim EmailIcon As String
                Dim storyName As String
                Dim storyOverview As String
                Dim storyBody As String
                Dim qs As String = ""
                Dim layout As New blockClass
                '
                Call layout.Load(newsBody)
                '
                PrinterIcon = "<img border=0 src=/ccLib/images/IconPrint.gif>"
                EmailIcon = "<img border=0 src=/ccLib/images/IconEmail.gif>"
                '
                If storyId = 0 Then
                    Call layout.SetClassInner("newsBodyStory", "<span class=""ccError"">The requested story is currently unavailable.</span>")
                Else
                    Call cs.Open(ContentNameNewsletterStories, "ID=" & storyId)
                    If cs.OK() Then
                        storyName = cs.GetText("name")
                        If isEditing Then
                            storyName = cs.GetEditLink() & storyName
                        End If
                        storyBody = cs.GetText("body")
                        storyOverview = cs.GetText("Overview")
                        If storyBody = "" Then
                            storyBody = storyOverview
                        End If
                        IssueID = cs.GetInteger("newsletterId")
                        '
                        returnHtml &= cs.GetEditLink()
                        If cs.GetBoolean("AllowPrinterPage") Then
                            qs = cp.Doc.RefreshQueryString
                            qs = cp.Utils.ModifyQueryString(qs, RequestNameStoryId, storyId.ToString())
                            qs = cp.Utils.ModifyQueryString(qs, RequestNameFormID, FormDetails.ToString())
                            qs = cp.Utils.ModifyQueryString(qs, "ccIPage", "l6d09a10sP")
                            returnHtml &= "<div class=""PrintIcon""><a target=_blank href=""?" & qs & """>" & PrinterIcon & "</a>&nbsp;<a target=_blank href=""" & qs & """><nobr>Printer Version</nobr></a></div>"
                        End If
                        If cs.GetBoolean("AllowEmailPage") Then
                            Link = "mailto:?SUBJECT=" & cp.Site.GetText("Email link subject", "A link to the " & cp.Site.DomainPrimary & " newsletter") & "&amp;BODY=http://" & cp.Site.DomainPrimary & cp.Site.AppRootPath & cp.Request.Page & Replace(refreshQueryString, "&", "%26") & RequestNameStoryId & "=" & storyId & "%26" & RequestNameFormID & "=" & FormDetails
                            returnHtml &= "<div class=""EmailIcon""><a target=_blank href=""?" & Link & """>" & EmailIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Email this page</nobr></a></div>"
                        End If
                        Call layout.SetClassInner("newsBodyCaption", storyName)
                        Call layout.SetClassInner("newsBodyStory", storyBody)
                        '
                        ' update RSS fields if empty
                        '
                        If Not isEditing Then
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
                                        Call cs.SetField("RSSDatePublish", PublishDate.ToString())
                                    End If
                                End If
                            End If
                            '
                            If (storyName <> "") And (cs.GetText("RSSTitle") = "") Then
                                rssChange = True
                                Call cs.SetField("RSSTitle", cs.GetText("name"))
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
                                    qs = refreshQueryString
                                    qs = cp.Utils.ModifyQueryString(qs, RequestNameStoryId, CStr(storyId))
                                    qs = cp.Utils.ModifyQueryString(qs, RequestNameFormID, FormDetails.ToString())
                                    qs = cp.Utils.ModifyQueryString(qs, "method", "")
                                    rssChange = True
                                    Call cs.SetField("RSSLink", Link & "?" & qs)
                                End If
                            End If
                            If rssChange Then
                                Call cp.Utils.ExecuteAddonAsProcess("RSS Feed Process")
                            End If
                        End If
                    End If
                    Call cs.Close()
                End If
                '
                returnHtml = layout.GetHtml()
            Catch ex As Exception
                Call handleError(cp, ex, "getNewsletterBodyDetails")
            End Try
            Return returnHtml
        End Function
        '
        Private Function template(x As Integer) As String
            Dim returnHtml As String = ""
            Try

            Catch ex As Exception
                'Call handleError(cp, ex, "template")
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
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterBodyClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace
