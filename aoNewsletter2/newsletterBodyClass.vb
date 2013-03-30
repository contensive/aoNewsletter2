Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterBodyClass
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
        Private Sub handleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterBodyClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
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

        Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
            'On Error GoTo ErrorTrap

            Csv = CsvObject

            Call Init(MainObject)

            Execute = GetContent(OptionString)

            'Exit Function
            'ErrorTrap:
            'Call HandleError("newsletterBodyClass", "Execute")
        End Function

        Public Sub Init(MainObject As Object)
            '
            Main = MainObject
            '
            Dim Common As New newsletterCommonClass
            '
            Call Common.UpgradeAddOn(Main)
            EncodeCopyNeeded = (Main.ContentServerVersion < "3.3.947")

            isManager = cp.User.IsContentManager("Newsletters")

            Exit Sub
            '
            'ErrorTrap:
            'Call HandleError(cp, ex, "Init")
        End Sub
        '
        Public Function GetContent(OptionString As String) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            Dim FileToField As Integer
            Dim DaFile As String
            Dim DaField As String
            Dim FileCounter As Integer
            Dim sql2 As String
            Dim sqlP As Integer
            Dim FileCount As Integer
            Dim FileId As Integer
            Dim Common As New newsletterCommonClass
            Dim NewsletterName As String
            Dim NewsletterProperty As String
            Dim Parts() As String
            '
            If Not (Main Is Nothing) Then
                '
                NewsletterName = cp.Doc.GetText("Newsletter", OptionString)
                archiveIssuesToDisplay = kmaEncodeInteger(cp.Doc.GetText("Archive Issues To Display", OptionString))
                '
                If NewsletterName <> "" Then
                    '
                    ' If newsletterNavClass used without PageClass, Newsletter is in the OptionString, Issue is in QS
                    '
                    NewsletterID = cp.Content.GetRecordID(ContentNameNewsletters, NewsletterName)
                    Call cp.Site.TestPoint("GetIssueID call 2, NewsletterID=" & NewsletterID)
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
                '
                WorkingQueryStringPlus = cp.Doc.RefreshQueryString
                '
                If WorkingQueryStringPlus = "" Then
                    WorkingQueryStringPlus = "?"
                Else
                    WorkingQueryStringPlus = "?" & WorkingQueryStringPlus & "&"
                End If
                '
                MonthSelected = cp.Doc.GetInteger(RequestNameMonthSelectd)
                YearSelected = cp.Doc.GetInteger(RequestNameYearSelected)
                ButtonValue = Main.GetStreamText("Button")
                SearchKeywords = Main.GetStreamText(RequestNameSearchKeywords)
                '
                RecordsPerPage = Main.GetSiteProperty("Newsletter Search Results Records Per Page", "3", True)
                '
                RecordTop = cp.Doc.GetInteger(RequestNameRecordTop)
                '
                PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
                If PageNumber = 0 Then
                    PageNumber = 1
                End If

                Stream = GetForm()
                GetContent = Stream
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetContent")
        End Function
        '
        Private Function GetForm() As String
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
                    Stream = Stream & GetArchiveList()
                Case FormDetails
                    Call cp.Site.TestPoint("GetForm Entering GetNewsletterBodyDetails")
                    Stream = Stream & GetNewsletterBodyDetails(IssuePageID)
                Case Else
                    Call cp.Site.TestPoint("GetForm Entering GetNewsletterBodyOverview")
                    FormID = FormIssue
                    Stream = Stream & GetNewsletterBodyOverview(IssueID, IssuePageID)
            End Select
            '    '
            '    Select Case ButtonValue
            '        Case FormButtonViewArchives
            ' '           Stream = Stream & GetArchiveList()
            '        Case FormButtonViewNewsLetter
            ' '           Stream = Stream & GetArchiveList()
            '    End Select
            '
            GetForm = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetForm")
        End Function
        '
        Private Function GetArchiveList() As String
            'On Error GoTo ErrorTrap
            '
            Dim CSPointer As CPCSBaseClass = cp.CSNew()
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
            Dim SearchPointer As Integer
            Dim SearchResult As String
            Dim SearchSQL As String
            '
            Dim NumberofPages As Integer
            Dim PageCount As Integer
            '
            Dim sql2 As String
            Dim SQLPointer As Integer
            Dim FileCount As Integer
            '
            Dim BriefCopyFileName As String
            '
            Dim RowCount As Integer
            Dim SQLCriteria As String
            '
            Dim YearsWanted As Integer
            Dim BlockSearchForm As Boolean
            '
            YearsWanted = kmaEncodeInteger(Main.GetSiteProperty("Newsletter years wanted", 1))
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
            sql2 = sql2 & " AND (NL.NewsletterID=" & Main.EncodeSQLNumber(NewsletterID) & ")"
            If MonthSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and month(nl.publishdate) = " & MonthSelected
            End If
            If YearSelected <> 0 Then
                ThisSQL2 = ThisSQL2 & " and year(nl.publishdate) = " & YearSelected
            End If
            If SearchKeywords <> "" Then
                sql2 = sql2 & " and ((nlp.Body like '%" & SearchKeywords & "%' )or (nlp.name  like '%" & SearchKeywords & "%') or (nlp.Overview  like '%" & SearchKeywords & "%'))"
            End If
            SQLPointer = Main.OpenCSSQL("Default", sql2)
            If Main.CSOK(SQLPointer) Then
                FileCount = cs.getInteger(SQLPointer, "count")
                NumberofPages = FileCount / RecordsPerPage
                If NumberofPages <> Int(NumberofPages) Then
                    NumberofPages = NumberofPages + 1
                    NumberofPages = Int(NumberofPages)
                End If
                If NumberofPages = 0 Then
                    NumberofPages = 1
                End If
            End If
            Call Main.CloseCS(SQLPointer)
            '
            Stream = Stream & Main.GetFormStart
            Stream = Stream & Main.GetFormInputHidden(RequestNameFormID, FormArchive)
            'Colors = "#ffffff"
            '
            '
            If (ButtonValue <> FormButtonViewNewsLetter) And (ButtonValue <> FormButtonViewArchives) Then
                '
                ' List a page of archive issues
                '
                If (MonthSelected = 0) And (YearSelected = 0) Then
                    'Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                    '
                    'ThisSQL = " SELECT  TOP 6 * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & Main.EncodeSQLNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    ThisSQL = " SELECT  TOP " & archiveIssuesToDisplay & " * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & Main.EncodeSQLNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    '
                    CSPointer = Main.OpenCSSQL("Default", ThisSQL)
                    If Main.CSOK(CSPointer) Then
                        'Stream = Stream & "<TR>"
                        Stream = Stream & Main.GetContentCopy2(PageNameArchives, , "<h2>Archive Issues</h2>")
                        'Stream = Stream & "<TD>" & Main.GetContentCopy2(PageNameArchives, , "<h2>Archive Issues</h2>") & "</TD>"
                        'Stream = Stream & "</TR>"
                        Do While Main.CSOK(CSPointer)
                            'Stream = Stream & "<TR>"
                            'Stream = Stream & "<TD BGCOLOR= """ & Colors & """>"
                            ' 1/1/09 - JK - always linked to the root path (approotpath)
                            Stream = Stream & Main.GetCSRecordEditLink(CSPointer) & "<a href=""" & WorkingQueryStringPlus & RequestNameIssueID & "=" & cs.getInteger(CSPointer, "ID") & """>" & GetIssuePublishDate(cs.getInteger(CSPointer, "ID")) & " " & Main.GetCSText(CSPointer, "Name") & "</a>"
                            'Stream = Stream & Main.GetCSRecordEditLink(CSPointer) & "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssueID & "=" & cs.getInteger(CSPointer, "ID") & """>" & GetIssuePublishDate(cs.getInteger(CSPointer, "ID")) & " " & Main.GetCSText(CSPointer, "Name") & "</a>"
                            Stream = Stream & "<br>"
                            If EncodeCopyNeeded Then
                                Stream = Stream & Main.EncodeContent(Main.GetCS(CSPointer, "Overview"), , True, False, False, True, True, True, True, "")
                                'Stream = Stream & Main.EncodeContent(Main.GetCS(CSPointer, "Overview"), , True, False, False, True, True, True, True, False) & "</TD>"
                            Else
                                Stream = Stream & Main.GetCS(CSPointer, "Overview")
                                'Stream = Stream & Main.GetCS(CSPointer, "Overview") & "</TD>"
                            End If
                            'Stream = Stream & "</TR>"
                            Call Main.NextCSRecord(CSPointer)
                            If Colors = "#ffffff" Then
                                Colors = "#E0E0E0"
                            Else
                                Colors = "#ffffff"
                            End If
                        Loop
                    Else
                        BlockSearchForm = True
                        'Stream = Stream & "<TR>"
                        Stream = Stream & "<span class=""ccError"">" & Main.GetSiteProperty(SitePropertyNoNewsletterArchives, "There are currently no archived issues.", True) & "</span>"
                        'Stream = Stream & "<TD><span class=""ccError"">" & Main.GetSiteProperty(SitePropertyNoNewsletterArchives, "There are currently no archived issues.", True) & "</span></TD>"
                        'Stream = Stream & "</TR>"
                    End If
                    Call Main.CloseCS(CSPointer)
                    'Stream = Stream & "</TABLE>"
                End If
            End If
            If ButtonValue = FormButtonViewArchives Then
                '
                ' List search results of archive issues
                '
                'Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
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
                SearchPointer = Main.OpenCSSQL("Default", ThisSQL2, RecordsPerPage, PageNumber)
                If Not Main.CSOK(SearchPointer) Then
                    Stream = Stream & Main.GetContentCopy2("Newsletter Search No Results Found", , "No results were found")
                Else
                    Stream = Stream & Main.GetContentCopy2("Newsletter Search Results Found", , "Search results")
                    Do While Main.CSOK(SearchPointer) And RowCount < RecordsPerPage
                        SearchResult = Main.GetCSText(SearchPointer, "nlpname")
                        Dim thisid As Integer
                        thisid = cs.getInteger(SearchPointer, "ID")
                        If Colors = "#E0E0E0" Then
                            Colors = "#ffffff"
                        Else
                            Colors = "#E0E0E0"
                        End If
                        'Stream = Stream & "<tr><td  BGCOLOR= """ & Colors & """ style=""border-top:1px solid #c0c0c0;padding:20px;"">"
                        Stream = Stream & "<div  class=""NewsletterBody"">"
                        Stream = Stream & "<div  class=""Headline"">" & SearchResult & "</div>"
                        BriefCopyFileName = Main.GetCSText(SearchPointer, "Overview")
                        If BriefCopyFileName = "" Then
                            If Main.GetCSBoolean(SearchPointer, "AllowReadMore") Then
                                BriefCopyFileName = Main.GetCS(SearchPointer, "Body")
                            Else
                                BriefCopyFileName = Main.GetContentCopy2("Newsletter Article Access Denied", , "You do not have access to this article")
                            End If
                        End If
                        Stream = Stream & "<div  class=""Overview"">" & Main.GetCS(SearchPointer, "Overview") & "</div>"
                        'Stream = Stream & Main.GetCSText(SearchPointer, "Overview")
                        'Stream = Stream & "<br><br>"
                        'Stream = Stream & "</tr></td>"
                        If (Main.GetCSBoolean(SearchPointer, "AllowReadMore")) And (Main.GetCS(SearchPointer, "Body") <> "") Then
                            Stream = Stream & "<a href=""" & WorkingQueryStringPlus & "formid=400&" & RequestNameIssuePageID & "=" & cs.getInteger(SearchPointer, "ThisID") & """>"
                            Stream = Stream & "Read More"
                            Stream = Stream & "</a>"
                        End If
                        Stream = Stream & "</div>"
                        Call Main.NextCSRecord(SearchPointer)
                        RowCount = RowCount + 1
                    Loop
                End If
                '
                If FileCount <> 0 Then
                    'Stream = Stream & "<tr><td align=center>"
                    Stream = Stream & "<div  class=""NewsletterBody""><div class=""GoToPageLine"">Go to Page&nbsp;&nbsp;"
                    Do While PageCount <= NumberofPages
                        'Stream = Stream & "<a href=""" & Main.ServerPage & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """> Page " & (PageCount) & "</a>"
                        ' 1/1/09 - JK - alays linked to root path
                        Stream = Stream & "<a href=""" & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """>" & (PageCount) & "</a>"
                        'Stream = Stream & "<a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """>" & (PageCount) & "</a>"
                        PageCount = PageCount + 1
                        Stream = Stream & "&nbsp;&nbsp;&nbsp;"
                    Loop
                    Stream = Stream & "</div></div>"
                End If
                'Stream = Stream & "</TABLE>"
            End If
            '
            If Not BlockSearchForm Then
                '
                ' Display search form
                '
                'Stream = Stream & "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                'Stream = Stream & "<tr>"
                'Stream = Stream & "<td>"
                Stream = Stream & Main.GetContentCopy2("Newsletter Search Copy", , "<h2>Archive Search</h2>")
                'Stream = Stream & "</td>"
                'Stream = Stream & "</tr>"
                '
                'Stream = Stream & "<tr>"
                'Stream = Stream & "<td>"
                ' 1 ** drop down select list 2007 Issues (all issues in 2007)
                Stream = Stream & "<div>" & Main.GetFormInputSelect(RequestNameIssueID, "", ContentNameNewsletterIssues, "(Publishdate<" & KmaEncodeSQLDate(Now) & ")AND(NewsletterID=" & Main.EncodeSQLNumber(NewsletterID) & ")") & " " & Main.GetFormButton(FormButtonViewNewsLetter) & "</div>"
                ' ** need a button to view the newsletter
                'Stream = Stream & "</td>"
                'Stream = Stream & "</tr>"
                '
                Stream = Stream & "<div>&nbsp;</div>"
                'Stream = Stream & "<tr>"
                'Stream = Stream & "<td>"
                Stream = Stream & "<div>keyword search<br>"
                Stream = Stream & Main.GetFormInputText(RequestNameSearchKeywords, , , 50) & "</div>"
                'Stream = Stream & "</td>"
                'Stream = Stream & "</tr>"
                '
                'Stream = Stream & "<tr>"
                'Stream = Stream & "<td>"
                MonthString = MonthString & "Month <select size=""1"" name=""" & RequestNameMonthSelectd & """>"
                MonthString = MonthString & "<option selected>Month</option>"
                For MonthCounter = 1 To 12
                    MonthString = MonthString & "<option "
                    MonthString = MonthString & "value=""" & MonthCounter & """>" & MonthName(MonthCounter) & "</option>"
                Next
                MonthString = MonthString & "</select> "
                '
                YearString = YearString & "Year <select size=""1"" name=""" & RequestNameYearSelected & """>"
                YearString = YearString & "<option selected>Year</option>"
                'For YearCounter = (Year(Now) - 5) To (Year(Now))
                For YearCounter = (Year(Now) - YearsWanted) To (Year(Now))
                    YearString = YearString & "<option "
                    YearString = YearString & "value=""" & YearCounter & """>" & YearCounter & "</option>"
                Next
                YearString = YearString & "</select>"
                Stream = Stream & "<div>&nbsp;</div>"
                Stream = Stream & "<div>" & MonthString & "&nbsp;&nbsp;&nbsp;" & YearString & "&nbsp;&nbsp;&nbsp;&nbsp;" & Main.GetFormButton(FormButtonViewArchives) & "</div>"
                'Stream = Stream & "</td>"
                'Stream = Stream & "</tr>"
                'Stream = Stream & "</TABLE>"
            End If
            '
            'Stream = Stream & "</TABLE>"
            Stream = Stream & Main.GetFormEnd
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
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD colspan=2 width=""60%"">" & Innards & "</TD>"
            Stream = Stream & "</TR>"
            '
            GetFormRow = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("DonationClass", "GetFormRow2")
        End Function
        '
Private Function GetSpacer(Optional Height As Integer, Optional Width As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            If Height = 0 Then
                Height = 1
            End If
            If Width = 0 Then
                Width = 1
            End If
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
Private Function GetArticleAccess(ArticleID As Integer, Optional GivenGroupID As Integer) As Boolean
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim AccessFlag As Boolean
            Dim ThisTest As String
            '
            If GivenGroupID <> 0 Then
                Call cs.open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                If Not Main.CSOK(cs) Then
                    GetArticleAccess = True
                Else
                    Do While Main.CSOK(cs)
                        If cs.getInteger("GroupID") = GivenGroupID Then
                            GetArticleAccess = True
                        End If
                        Call Main.NextCSRecord(cs)
                    Loop
                End If
                Call cs.close()
            Else
                If Not isManager Then
                    Call cs.open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                    If Not Main.CSOK(cs) Then
                        GetArticleAccess = True
                    Else
                        Do While Main.CSOK(cs)
                            ThisTest = Main.GetCSLookup(cs, "GroupID")
                            '
                            '
                            If ThisTest <> "" Then
                                If Main.IsGroupMember(ThisTest) Then
                                    GetArticleAccess = True
                                End If
                            End If
                            Call Main.NextCSRecord(cs)
                        Loop
                    End If
                    Call cs.close()
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
        Private Function GetIssuePublishDate(IssueID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim CSPointer As CPCSBaseClass = cp.CSNew()
            Dim IssueDate As String
            Dim Stream As String
            '
            CSPointer = Main.OpenCSContent(ContentNameNewsletterIssues, "ID=" & IssueID, , , "PublishDate")
            If Main.CSOK(CSPointer) Then
                IssueDate = Main.GetCSDate(CSPointer, "PublishDate")
                If IsDate(IssueDate) Then
                    Stream = MonthName(Month(IssueDate), True) & " " & Day(IssueDate) & ", " & Year(IssueDate)
                End If
            End If
            Call Main.CloseCS(CSPointer)
            '
            '
            GetIssuePublishDate = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetArticleAccess")
        End Function
        '
        Private Function GetEmailBody(TemplateCopy As String, LocalGroupID As Integer) As String
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
                            Stream = Stream & Replace(InnerValue, TemplateReplacementBody, GetNewsletterBodyOverview(IssueID, IssuePageID, LocalGroupID))
                        Case TemplateReplacementNav
                            Call Navigation.Init(Main)
                            Stream = Stream & Replace(InnerValue, TemplateReplacementNav, Navigation.GetContent("NavigationLayout=Vertical", LocalGroupID))
                            'Case TemplateReplacementTitle
                            '    Call Mast.Init(Main)
                            '    Stream = Stream & Replace(InnerValue, TemplateReplacementTitle, Mast.GetContent("MastMode=TitleOnly"))
                            'Case TemplateReplacementPubDate
                            '    Call Mast.Init(Main)
                            '    Stream = Stream & Replace(InnerValue, TemplateReplacementPubDate, Mast.GetContent("MastMode=PubDateOnly"))
                    End Select
                    Stream = Stream & InnerTemplateArray(1)
                Else
                    Stream = Stream & TemplateArray(TemplateArrayPointer)
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
        Private Function GetOverview(PageID As Integer) As String
            '
            Dim cs As cpcsBaseClass = cp.csNew()
            '
            Call cs.open(ContentNameNewsletterIssuePages, "ID=" & Main.EncodeSQLNumber(PageID), , , , , "Overview")
            If Main.CSOK(CS) Then
                GetOverview = cs.getText("Overview")
            End If
            Call cs.close()
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("Newsletter.newsletterBodyClass", "GetOverview")
        End Function
        '
Friend Function GetNewsletterBodyOverview(IssueID As Integer, IssuePageID As Integer, Optional GivenGroupID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim AddLink As String
            Dim Controls As String
            Dim IssueSQL As String
            Dim IssuePointer As Integer
            Dim NewIssueId As Integer
            Dim MaxIssueID As Integer
            Dim Stream As String
            Dim cs As cpcsBaseClass = cp.csNew()
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
            Dim Common As New newsletterCommonClass
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
                Stream = Stream & Main.GetCSRecordEditLink(cs) & Main.GetCS(cs, "Cover")
            End If
            Call cs.close()
            '
            If IssuePageID <> 0 Then
                Criteria = ""
                MainSQL = "" _
                    & " select p.categoryId,c.name as CategoryName" _
                    & " from NewsletterIssuePages p" _
                    & " left join NewsletterIssueCategories c on c.id=p.categoryId" _
                    & " where (p.ID=" & Main.EncodeSQLNumber(IssuePageID) & ")" _
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
                'CS = Main.OpenCSSQL("Default", MainSQL)
                '
            End If
            Call cp.Site.TestPoint("MainSQL: " & MainSQL)
            CS = Main.OpenCSSQL("Default", MainSQL)
            '
            Stream = Stream & vbCrLf & "<!-- Start NewsletterBody -->" & vbCrLf
            Stream = Stream & "<div class=""NewsletterBody"">"
            '
            If Main.CSOK(CS) Then
                Do While Main.CSOK(CS)
                    CategoryID = cs.getInteger("CategoryID")
                    CategoryName = cs.getText("CategoryName")
                    '
                    CS2 = Main.OpenCSContent(ContentNameNewsletterIssuePages, "(CategoryID=" & CategoryID & ") AND (NewsletterID=" & IssueID & ")", "SortOrder")
                    If Main.IsCSOK(CS2) Then
                        '
                        ' there are stories under this topic, wrap in div to allow a story indent
                        '
                        Stream = Stream & vbCrLf & "<div class=""NewsletterTopic"">"
                        If RecordCount <> 0 Then
                            If Main.IsLinkAuthoring(ContentNameNewsletterIssuePages) Then
                                ' 1/1/09 JK added hint wrapper
                                Stream = Stream & Main.GetAdminHintWrapper("<a href=""" & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssueID & "=" & IssueID & "&" & RequestNameSortUp & "=" & CategoryID & """>[Move Up]</a> ")
                            End If
                        End If
                        Stream = Stream & CategoryName
                        Stream = Stream & "</div>"
                        '
                        Stream = Stream & vbCrLf & "<div class=""NewsletterTopicStory"">"
                        Do While Main.CSOK(CS2)
                            Stream = Stream & GetStoryOverview(CS2)
                            Call Main.NextCSRecord(CS2)
                        Loop
                        Stream = Stream & "</div>"
                    End If
                    Call Main.CloseCS(CS2)
                    Call Main.NextCSRecord(CS)
                    RecordCount = RecordCount + 1
                Loop
            End If
            '
            Call cs.close()
            '
            Stream = Stream & GetUnrelatedStories(IssuePageID)
            '
            IssueSQL = " Select max(id) as MaxIssueID from newsletterissues"
            IssuePointer = Main.OpenCSSQL("Default", IssueSQL)
            If Main.CSOK(IssuePointer) Then
                MaxIssueID = cs.getInteger(IssuePointer, "maxissueid")
            End If
            NewIssueId = MaxIssueID + 1
            Call Main.CloseCS(IssuePointer)
            'Controls = Common.GetAuthoringLinks(cp, IssuePageID, IssueID, NewsletterID, WorkingQueryStringPlus)
            'If Controls <> "" Then
            '    Stream = Stream & Main.GetAdminHintWrapper(Controls)
            'End If
            '    If Main.IsLinkAuthoring(ContentNameNewsletterIssues) Then
            '        ' 1/1/09 added admin hint wrapper
            '        Controls = "<br /><br />"
            '        If IssuePageID <> 0 Then
            '            Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.site.getText( "adminUrl" ) & "?cid=" & cp.Content.getid(ContentNameNewsletterIssuePages) & "&af=4&id=" & IssuePageID & "&" & ReferLink & """>Edit this page</a></div>"
            '        End If
            '        If IssueID <> 0 Then
            '            Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.site.getText( "adminUrl" ) & "?cid=" & cp.Content.getid(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div>"
            '            Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.site.getText( "adminUrl" ) & "?cid=" & cp.Content.getid(ContentNameNewsletterIssuePages) & "&af=4&aa=2&ad=1&wc=NewsletterID=" & IssueID & "&" & ReferLink & """>Add a new page to this issue</a></div>"
            '        End If
            '        Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & cp.Site.DomainPrimary & cp.site.getText( "adminUrl" ) & "?cid=" & cp.Content.getid(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div>"
            '        Stream = Stream & Main.GetAdminHintWrapper(Controls)
            '    End If
            '
            Stream = Stream & "</div>"
            '
            Stream = Stream & Main.GetRecordAddLink(ContentNameNewsletterIssuePages, "Newsletterid=" & IssueID)
            '
            Stream = Stream & vbCrLf & "<!-- End NewsletterBody -->" & vbCrLf
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
                If Main.CSOK(CS) Then
                    Caption = Main.GetSiteProperty("Newsletter Caption Other Stories", "")
                    If Caption <> "" Then
                        Stream = Stream & vbCrLf & "<div class=""NewsletterTopic"">" & Caption & "</div>"
                    End If
                    Do While Main.CSOK(CS)
                        Stream = Stream & GetStoryOverview(CS)
                        Call Main.NextCSRecord(CS)
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
        Private Function GetStoryOverview(CS As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim StoryID As Integer
            Dim StoryAccessString As String
            Dim Stream As String
            Dim Common As New newsletterCommonClass
            Dim storyBookmark As String
            '
            StoryID = cs.getInteger("ID")
            storyBookmark = "story" & StoryID
            StoryAccessString = Common.GetArticleAccessString(cp, StoryID)
            '
            If StoryAccessString <> "" Then
                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & StoryAccessString & """>"
            End If
            '
            Stream = Stream & vbCrLf & "<div class=""Headline"" id=""" & storyBookmark & """>"
            If FormID <> FormEmail Then
                Stream = Stream & Main.GetCSRecordEditLink(CS)
            End If
            Stream = Stream & cs.getText("Name") & "</div>"
            'Stream = Stream & "<a name=""" & storyBookmark & """>" &cs.getText( "Name") & "&nbsp;" & "</a></div>"
            If IssuePageID <> 0 Then
                Call Main.TrackContent(ContentNameNewsletterIssuePages, IssuePageID)
                If EncodeCopyNeeded Then
                    Stream = Stream & "<div class=""Copy"">" & Main.EncodeContent(cs.getText("Body "), , True, False, False, True, True, True, True, "") & "</div>"
                Else
                    Stream = Stream & "<div class=""Copy"">" & cs.getText("Body ") & "</div>"
                End If
            Else
                Stream = Stream & "<div class=""Overview"">"
                Stream = Stream & cs.getText("Overview")
                If Main.GetCSBoolean(CS, "AllowReadMore") Then
                    ' 1/1/09 JK - always linked to root path
                    Stream = Stream & "<div class=""ReadMore""><a href=""" & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & cs.getInteger("ID") & "&" & RequestNameFormID & "=" & FormDetails & """>Read More</a></div>"
                    'Stream = Stream & "<div class=""ReadMore""><a href=""http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & cs.getInteger("ID") & "&" & RequestNameFormID & "=" & FormDetails & """>Read More</a></div>"
                    'Else
                    '    Stream = Stream & "<div class=""ReadMore"">&nbsp;</div>"
                End If
                Stream = Stream & "</div>"
            End If
            If StoryAccessString <> "" Then
                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >"
            End If
            '
            GetStoryOverview = Stream
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetStoryOverview")
        End Function
        '
        Private Function GetNewsletterBodyDetails(IssuePageID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim rssChange As Boolean
            Dim expirationDate As Date
            Dim PublishDate As Date
            Dim Pos As Integer
            Dim recordDate As Date
            Dim CSIssue As Integer
            Dim Copy As String
            Dim Stream As String
            Dim cs As cpcsBaseClass = cp.csNew()
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
                Call cs.open(ContentNameNewsletterIssuePages, "ID=" & IssuePageID)
                If Main.CSOK(CS) Then
                    storyName = cs.getText("name")
                    storyOverview = cs.getText("Overview")
                    IssueID = cs.getInteger("newsletterId")
                    '
                    Stream = Stream & Main.GetCSRecordEditLink(CS)
                    If Main.GetCSBoolean(CS, "AllowPrinterPage") Then
                        Link = WorkingQueryStringPlus & RequestNameIssuePageID & "=" & IssuePageID & "&" & RequestNameFormID & "=" & FormDetails & "&ccIPage=l6d09a10sP"
                        Stream = Stream & "<div class=""PrintIcon""><a target=_blank href=""" & Link & """>" & PrinterIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Printer Version</nobr></a></div>"
                    End If
                    If Main.GetCSBoolean(CS, "AllowEmailPage") Then
                        Link = "mailto:?SUBJECT=" & Main.GetSiteProperty("Email link subject", "A link to the " & cp.Site.DomainPrimary & " newsletter", True) & "&amp;BODY=http://" & cp.Site.DomainPrimary & Main.ServerAppRootPath & Main.ServerPage & Replace(WorkingQueryStringPlus, "&", "%26") & RequestNameIssuePageID & "=" & IssuePageID & "%26" & RequestNameFormID & "=" & FormDetails
                        Stream = Stream & "<div class=""EmailIcon""><a target=_blank href=""" & Link & """>" & EmailIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Email this page</nobr></a></div>"
                    End If
                    Stream = Stream & "<div class=""NewsletterBody"">"
                    Stream = Stream & "<div class=""Headline"">" & cs.getText("Name") & "</div>"
                    Stream = Stream & "<div class=""Copy"">" & cs.getText("Body") & "</div>"
                    Stream = Stream & "</div>"
                    '
                    ' update RSS fields if empty
                    '
                    rssChange = False
                    If (IssueID <> 0) Then
                        If (Main.GetCSDate(CS, "RSSDatePublish") = CDate(0)) Then
                            CSIssue = Main.OpenCSContent(ContentNameNewsletterIssues, "id=" & KmaEncodeSQLNumber(IssueID))
                            If Main.IsCSOK(CSIssue) Then
                                PublishDate = Main.GetCSDate(CSIssue, "publishDate")
                            End If
                            Call Main.CloseCS(CSIssue)
                            If (PublishDate <> CDate(0)) Then
                                rssChange = True
                                Call Main.SetCS(CS, "RSSDatePublish", PublishDate)
                            End If
                        End If
                    End If
                    '
                    If (storyName <> "") And (cs.getText("RSSTitle") = "") Then
                        rssChange = True
                        Call Main.SetCS(CS, "RSSTitle", storyName)
                    End If
                    '
                    If (storyOverview <> "") And (cs.getText("RSSDescription") = "") Then
                        rssChange = True
                        Copy = Main.ConvertHTML2Text(storyOverview)
                        Call Main.SetCS(CS, "RSSDescription", Copy)
                    End If
                    '
                    If (cs.getText("RSSLink") = "") Then
                        Link = Main.ServerLink
                        If InStr(1, Link, cp.site.getText("adminUrl"), vbTextCompare) = 0 Then
                            Pos = InStr(1, Link, "?")
                            If Pos > 0 Then
                                Link = Left(Link, Pos - 1)
                            End If
                            Copy = cp.Doc.RefreshQueryString()
                            Copy = cp.Utils.ModifyQueryString(Copy, RequestNameIssuePageID, CStr(IssuePageID))
                            Copy = cp.Utils.ModifyQueryString(Copy, RequestNameFormID, FormDetails)
                            Copy = cp.Utils.ModifyQueryString(Copy, "method", "")
                            rssChange = True
                            Call Main.SetCS(CS, "RSSLink", Link & "?" & Copy)
                        End If
                    End If
                    If rssChange Then
                        Call Main.ExecuteAddonAsProcess("RSS Feed Process")
                    End If
                End If
                Call cs.close()
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
