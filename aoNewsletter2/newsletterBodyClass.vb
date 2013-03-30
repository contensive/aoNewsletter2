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
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
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
        Private PageNumber As Long
        '
        Private FormID As Long
        Private RecordsPerPage As Long
        '
        Private IssueID As Long
        Private IssuePageID As Long
        Private MonthSelected As Long
        Private YearSelected As Long
        Private ButtonValue As String
        Private RecordTop As Long
        Private SearchKeywords As String

        Private isManager As Boolean

        Private NewsletterID As Long
        Private archiveIssuesToDisplay As Long
        '
        'Private Main As MainClass
        Private Main As Object
        Private EncodeCopyNeeded As Boolean
        Private Csv As Object

        Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
            On Error GoTo ErrorTrap

            Csv = CsvObject

            Call Init(MainObject)

            Execute = GetContent(OptionString)

            Exit Function
ErrorTrap:
            Call HandleError("BodyClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
        End Function

        Public Sub Init(MainObject As Object)
            '
            Main = MainObject
            '
            Dim Common As New CommonClass
            '
            Call Common.UpgradeAddOn(Main)
            EncodeCopyNeeded = (Main.ContentServerVersion < "3.3.947")

            isManager = Main.IsContentManager("Newsletters")

            Exit Sub
            '
ErrorTrap:
            Call HandleError("NewsLetter", "Init", Err.Number, Err.Source, Err.Description, True, False)
        End Sub
        '
        Public Function GetContent(OptionString As String) As String
            On Error GoTo ErrorTrap
            '
            Dim Stream As String
            Dim FileToField As Long
            Dim DaFile As String
            Dim DaField As String
            Dim FileCounter As Long
            Dim sql2 As String
            Dim sqlP As Long
            Dim FileCount As Long
            Dim FileId As Long
            Dim Common As New CommonClass
            Dim NewsletterName As String
            Dim NewsletterProperty As String
            Dim Parts() As String
            '
            If Not (Main Is Nothing) Then
                '
                NewsletterName = Main.GetAddonOption("Newsletter", OptionString)
                archiveIssuesToDisplay = kmaEncodeInteger(Main.GetAddonOption("Archive Issues To Display", OptionString))
                '
                If NewsletterName <> "" Then
                    '
                    ' If NavClass used without PageClass, Newsletter is in the OptionString, Issue is in QS
                    '
                    NewsletterID = Main.GetRecordID(ContentNameNewsletters, NewsletterName)
                    Call Main.TestPoint("GetIssueID call 2, NewsletterID=" & NewsletterID)
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
                '
                WorkingQueryStringPlus = Main.RefreshQueryString
                '
                If WorkingQueryStringPlus = "" Then
                    WorkingQueryStringPlus = "?"
                Else
                    WorkingQueryStringPlus = "?" & WorkingQueryStringPlus & "&"
                End If
                '
                MonthSelected = Main.GetStreamInteger(RequestNameMonthSelectd)
                YearSelected = Main.GetStreamInteger(RequestNameYearSelected)
                ButtonValue = Main.GetStreamText("Button")
                SearchKeywords = Main.GetStreamText(RequestNameSearchKeywords)
                '
                RecordsPerPage = Main.GetSiteProperty("Newsletter Search Results Records Per Page", "3", True)
                '
                RecordTop = Main.GetStreamInteger(RequestNameRecordTop)
                '
                PageNumber = Main.GetStreamInteger(RequestNamePageNumber)
                If PageNumber = 0 Then
                    PageNumber = 1
                End If

                Stream = GetForm()
                GetContent = Stream
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetForm() As String
            On Error GoTo ErrorTrap
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
                    Call Main.TestPoint("GetForm Entering GetNewsletterBodyDetails")
                    Stream = Stream & GetNewsletterBodyDetails(IssuePageID)
                Case Else
                    Call Main.TestPoint("GetForm Entering GetNewsletterBodyOverview")
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
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetArchiveList() As String
            On Error GoTo ErrorTrap
            '
            Dim CSPointer As Long
            Dim Stream As String
            Dim Colors As String
            Dim ThisSQL As String
            Dim ThisSQL2 As String
            Dim MonthString As String
            Dim YearString As String
            Dim MonthCounter As Long
            Dim YearCounter As Long
            '
            Dim SelectedIssuePointer As Long
            Dim SelectedIssue As String
            '
            Dim SearchPointer As Long
            Dim SearchResult As String
            Dim SearchSQL As String
            '
            Dim NumberofPages As Long
            Dim PageCount As Long
            '
            Dim sql2 As String
            Dim SQLPointer As Long
            Dim FileCount As Long
            '
            Dim BriefCopyFileName As String
            '
            Dim RowCount As Long
            Dim SQLCriteria As String
            '
            Dim YearsWanted As Long
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
                FileCount = Main.GetCSInteger(SQLPointer, "count")
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
                            Stream = Stream & Main.GetCSRecordEditLink(CSPointer) & "<a href=""" & WorkingQueryStringPlus & RequestNameIssueID & "=" & Main.GetCSInteger(CSPointer, "ID") & """>" & GetIssuePublishDate(Main.GetCSInteger(CSPointer, "ID")) & " " & Main.GetCSText(CSPointer, "Name") & "</a>"
                            'Stream = Stream & Main.GetCSRecordEditLink(CSPointer) & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssueID & "=" & Main.GetCSInteger(CSPointer, "ID") & """>" & GetIssuePublishDate(Main.GetCSInteger(CSPointer, "ID")) & " " & Main.GetCSText(CSPointer, "Name") & "</a>"
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
                        Dim thisid As Long
                        thisid = Main.GetCSInteger(SearchPointer, "ID")
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
                            Stream = Stream & "<a href=""" & WorkingQueryStringPlus & "formid=400&" & RequestNameIssuePageID & "=" & Main.GetCSInteger(SearchPointer, "ThisID") & """>"
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
                        'Stream = Stream & "<a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameButtonValue & "=" & FormButtonViewArchives & "&" & RequestNamePageNumber & "=" & PageCount & "&" & RequestNameSearchKeywords & "=" & SearchKeywords & """>" & (PageCount) & "</a>"
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
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetArchiveList", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetFormRow(Innards As String) As String
            On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            Stream = Stream & "<TR>"
            Stream = Stream & "<TD colspan=2 width=""60%"">" & Innards & "</TD>"
            Stream = Stream & "</TR>"
            '
            GetFormRow = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("DonationClass", "GetFormRow2", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
Private Function GetSpacer(Optional Height As Long, Optional Width As Long) As String
            On Error GoTo ErrorTrap
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
            Exit Function
ErrorTrap:
            Call HandleError("LeftSideNavigation", "GetSpacer", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
Private Function GetArticleAccess(ArticleID As Long, Optional GivenGroupID As Long) As Boolean
            On Error GoTo ErrorTrap
            '
            Dim CSPointer As Long
            Dim AccessFlag As Boolean
            Dim ThisTest As String
            '
            If GivenGroupID <> 0 Then
                CSPointer = Main.OpenCSContent(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                If Not Main.CSOK(CSPointer) Then
                    GetArticleAccess = True
                Else
                    Do While Main.CSOK(CSPointer)
                        If Main.GetCSInteger(CSPointer, "GroupID") = GivenGroupID Then
                            GetArticleAccess = True
                        End If
                        Call Main.NextCSRecord(CSPointer)
                    Loop
                End If
                Call Main.CloseCS(CSPointer)
            Else
                If Not isManager Then
                    CSPointer = Main.OpenCSContent(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                    If Not Main.CSOK(CSPointer) Then
                        GetArticleAccess = True
                    Else
                        Do While Main.CSOK(CSPointer)
                            ThisTest = Main.GetCSLookup(CSPointer, "GroupID")
                            '
                            '
                            If ThisTest <> "" Then
                                If Main.IsGroupMember(ThisTest) Then
                                    GetArticleAccess = True
                                End If
                            End If
                            Call Main.NextCSRecord(CSPointer)
                        Loop
                    End If
                    Call Main.CloseCS(CSPointer)
                Else
                    GetArticleAccess = True
                End If
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "GetArticleAccess", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetIssuePublishDate(IssueID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim CSPointer As Long
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
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "GetArticleAccess", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetEmailBody(TemplateCopy As String, LocalGroupID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim Stream As String
            '
            Dim TemplateArray() As String
            Dim TemplateArrayCount As Long
            Dim TemplateArrayPointer As Long
            '
            Dim InnerTemplateArray() As String
            Dim InnerTemplateArrayCount As Long
            Dim InnerTemplateArrayPointer As Long
            '
            Dim InnerValue As String
            '
            Dim Navigation As New NavClass
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
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "GetEmailBody", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetOverview(PageID As Long) As String
            '
            Dim CS As Long
            '
            CS = Main.OpenCSContent(ContentNameNewsletterIssuePages, "ID=" & Main.EncodeSQLNumber(PageID), , , , , "Overview")
            If Main.CSOK(CS) Then
                GetOverview = Main.GetCSText(CS, "Overview")
            End If
            Call Main.CloseCS(CS)
            '
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter.BodyClass", "GetOverview", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
Friend Function GetNewsletterBodyOverview(IssueID As Long, IssuePageID As Long, Optional GivenGroupID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim AddLink As String
            Dim Controls As String
            Dim IssueSQL As String
            Dim IssuePointer As Long
            Dim NewIssueId As Long
            Dim MaxIssueID As Long
            Dim Stream As String
            Dim CS As Long
            Dim Criteria As String
            Dim Link As String
            Dim HasArticleAccess As Boolean
            Dim SQL As String
            Dim TableList As String
            Dim CSTopics As Long
            Dim MainSQL As String
            Dim PreviousCategoryName As String
            Dim CategoryName As String
            Dim RecordCount As Long
            Dim Common As New CommonClass
            Dim AccessString As String
            Dim StoryID As Long
            Dim StoryAccessString As String
            Dim Caption As String
            Dim FetchFlag As Boolean
            '
            Dim CategoryID As Long
            Dim CS2 As Long
            '
            TableList = "NewsletterIssuePages "
            '
            CS = Main.OpenCSContentRecord("Newsletter Issues", IssueID)
            If Main.IsCSOK(CS) Then
                Stream = Stream & Main.GetCSRecordEditLink(CS) & Main.GetCS(CS, "Cover")
            End If
            Call Main.CloseCS(CS)
            '
            If IssuePageID <> 0 Then
                Criteria = ""
                MainSQL = "" _
                    & " select p.categoryId,c.name as CategoryName" _
                    & " from NewsletterIssuePages p" _
                    & " left join NewsletterIssueCategories c on c.id=p.categoryId" _
                    & " where (p.ID=" & Main.EncodeSQLNumber(IssuePageID) & ")" _
                    & ""
                'CS = Main.OpenCSContent(ContentNameNewsletterIssuePages, Criteria, "SortOrder,DateAdded")
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
                'Call Main.TestPoint("MainSQL: " & MainSQL)
                'CS = Main.OpenCSSQL("Default", MainSQL)
                '
            End If
            Call Main.TestPoint("MainSQL: " & MainSQL)
            CS = Main.OpenCSSQL("Default", MainSQL)
            '
            Stream = Stream & vbCrLf & "<!-- Start NewsletterBody -->" & vbCrLf
            Stream = Stream & "<div class=""NewsletterBody"">"
            '
            If Main.CSOK(CS) Then
                Do While Main.CSOK(CS)
                    CategoryID = Main.GetCSInteger(CS, "CategoryID")
                    CategoryName = Main.GetCSText(CS, "CategoryName")
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
            Call Main.CloseCS(CS)
            '
            Stream = Stream & GetUnrelatedStories(IssuePageID)
            '
            IssueSQL = " Select max(id) as MaxIssueID from newsletterissues"
            IssuePointer = Main.OpenCSSQL("Default", IssueSQL)
            If Main.CSOK(IssuePointer) Then
                MaxIssueID = Main.GetCSInteger(IssuePointer, "maxissueid")
            End If
            NewIssueId = MaxIssueID + 1
            Call Main.CloseCS(IssuePointer)
            'Controls = Common.GetAuthoringLinks(Main, IssuePageID, IssueID, NewsletterID, WorkingQueryStringPlus)
            'If Controls <> "" Then
            '    Stream = Stream & Main.GetAdminHintWrapper(Controls)
            'End If
            '    If Main.IsLinkAuthoring(ContentNameNewsletterIssues) Then
            '        ' 1/1/09 added admin hint wrapper
            '        Controls = "<br /><br />"
            '        If IssuePageID <> 0 Then
            '            Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssuePages) & "&af=4&id=" & IssuePageID & "&" & ReferLink & """>Edit this page</a></div>"
            '        End If
            '        If IssueID <> 0 Then
            '            Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div>"
            '            Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssuePages) & "&af=4&aa=2&ad=1&wc=NewsletterID=" & IssueID & "&" & ReferLink & """>Add a new page to this issue</a></div>"
            '        End If
            '        Controls = Controls & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div>"
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
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetNewsletterBodyOverview", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '

        Private Function GetUnrelatedStories(IssuePageID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim Criteria As String
            Dim CS As Long
            Dim Caption As String
            Dim Stream As String
            '
            If IssuePageID = 0 Then
                Criteria = "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")"
                CS = Main.OpenCSContent(ContentNameNewsletterIssuePages, Criteria, "SortOrder,DateAdded")
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
                Call Main.CloseCS(CS)
            End If
            '
            GetUnrelatedStories = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "GetUnrelatedStories", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetStoryOverview(CS As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim StoryID As Long
            Dim StoryAccessString As String
            Dim Stream As String
            Dim Common As New CommonClass
            Dim storyBookmark As String
            '
            StoryID = Main.GetCSInteger(CS, "ID")
            storyBookmark = "story" & StoryID
            StoryAccessString = Common.GetArticleAccessString(Main, StoryID)
            '
            If StoryAccessString <> "" Then
                Stream = Stream & "<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & StoryAccessString & """>"
            End If
            '
            Stream = Stream & vbCrLf & "<div class=""Headline"" id=""" & storyBookmark & """>"
            If FormID <> FormEmail Then
                Stream = Stream & Main.GetCSRecordEditLink(CS)
            End If
            Stream = Stream & Main.GetCSText(CS, "Name") & "</div>"
            'Stream = Stream & "<a name=""" & storyBookmark & """>" & Main.GetCSText(CS, "Name") & "&nbsp;" & "</a></div>"
            If IssuePageID <> 0 Then
                Call Main.TrackContent(ContentNameNewsletterIssuePages, IssuePageID)
                If EncodeCopyNeeded Then
                    Stream = Stream & "<div class=""Copy"">" & Main.EncodeContent(Main.GetCSText(CS, "Body "), , True, False, False, True, True, True, True, "") & "</div>"
                Else
                    Stream = Stream & "<div class=""Copy"">" & Main.GetCSText(CS, "Body ") & "</div>"
                End If
            Else
                Stream = Stream & "<div class=""Overview"">"
                Stream = Stream & Main.GetCSText(CS, "Overview")
                If Main.GetCSBoolean(CS, "AllowReadMore") Then
                    ' 1/1/09 JK - always linked to root path
                    Stream = Stream & "<div class=""ReadMore""><a href=""" & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & Main.GetCSInteger(CS, "ID") & "&" & RequestNameFormID & "=" & FormDetails & """>Read More</a></div>"
                    'Stream = Stream & "<div class=""ReadMore""><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssuePageID & "=" & Main.GetCSInteger(CS, "ID") & "&" & RequestNameFormID & "=" & FormDetails & """>Read More</a></div>"
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
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "GetStoryOverview", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetNewsletterBodyDetails(IssuePageID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim rssChange As Boolean
            Dim expirationDate As Date
            Dim PublishDate As Date
            Dim Pos As Long
            Dim recordDate As Date
            Dim CSIssue As Long
            Dim Copy As String
            Dim Stream As String
            Dim CS As Long
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
                CS = Main.OpenCSContent(ContentNameNewsletterIssuePages, "ID=" & IssuePageID)
                If Main.CSOK(CS) Then
                    storyName = Main.GetCSText(CS, "name")
                    storyOverview = Main.GetCSText(CS, "Overview")
                    IssueID = Main.GetCSInteger(CS, "newsletterId")
                    '
                    Stream = Stream & Main.GetCSRecordEditLink(CS)
                    If Main.GetCSBoolean(CS, "AllowPrinterPage") Then
                        Link = WorkingQueryStringPlus & RequestNameIssuePageID & "=" & IssuePageID & "&" & RequestNameFormID & "=" & FormDetails & "&ccIPage=l6d09a10sP"
                        Stream = Stream & "<div class=""PrintIcon""><a target=_blank href=""" & Link & """>" & PrinterIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Printer Version</nobr></a></div>"
                    End If
                    If Main.GetCSBoolean(CS, "AllowEmailPage") Then
                        Link = "mailto:?SUBJECT=" & Main.GetSiteProperty("Email link subject", "A link to the " & Main.ServerHost & " newsletter", True) & "&amp;BODY=http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & Replace(WorkingQueryStringPlus, "&", "%26") & RequestNameIssuePageID & "=" & IssuePageID & "%26" & RequestNameFormID & "=" & FormDetails
                        Stream = Stream & "<div class=""EmailIcon""><a target=_blank href=""" & Link & """>" & EmailIcon & "</a>&nbsp;<a target=_blank href=""" & Link & """><nobr>Email this page</nobr></a></div>"
                    End If
                    Stream = Stream & "<div class=""NewsletterBody"">"
                    Stream = Stream & "<div class=""Headline"">" & Main.GetCSText(CS, "Name") & "</div>"
                    Stream = Stream & "<div class=""Copy"">" & Main.GetCSText(CS, "Body") & "</div>"
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
                    If (storyName <> "") And (Main.GetCSText(CS, "RSSTitle") = "") Then
                        rssChange = True
                        Call Main.SetCS(CS, "RSSTitle", storyName)
                    End If
                    '
                    If (storyOverview <> "") And (Main.GetCSText(CS, "RSSDescription") = "") Then
                        rssChange = True
                        Copy = Main.ConvertHTML2Text(storyOverview)
                        Call Main.SetCS(CS, "RSSDescription", Copy)
                    End If
                    '
                    If (Main.GetCSText(CS, "RSSLink") = "") Then
                        Link = Main.ServerLink
                        If InStr(1, Link, Main.SiteProperty_AdminURL, vbTextCompare) = 0 Then
                            Pos = InStr(1, Link, "?")
                            If Pos > 0 Then
                                Link = Left(Link, Pos - 1)
                            End If
                            Copy = Main.RefreshQueryString()
                            Copy = ModifyQueryString(Copy, RequestNameIssuePageID, CStr(IssuePageID))
                            Copy = ModifyQueryString(Copy, RequestNameFormID, FormDetails)
                            Copy = ModifyQueryString(Copy, "method", "")
                            rssChange = True
                            Call Main.SetCS(CS, "RSSLink", Link & "?" & Copy)
                        End If
                    End If
                    If rssChange Then
                        Call Main.ExecuteAddonAsProcess("RSS Feed Process")
                    End If
                End If
                Call Main.CloseCS(CS)
            End If
            '
            GetNewsletterBodyDetails = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("Newsletter", "GetNewsletterBodyDetails", Err.Number, Err.Source, Err.Description, True, False)
        End Function
    End Class
End Namespace
