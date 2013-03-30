
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterCommonClass
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
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterCommonClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Friend Function GetIssueID(Main As Object, NewsletterID As Long) As Long
            On Error GoTo ErrorTrap
            '
            Dim IssueID As Long
            '
            IssueID = Main.GetStreamInteger(RequestNameIssueID)
            '
            Call Main.TestPoint("GetIssueID - IssueID From Stream: " & IssueID)
            Call Main.TestPoint("GetIssueID - NewsletterID: " & NewsletterID)
            '
            If IssueID = 0 Then
                IssueID = GetCurrentIssueID(Main, NewsletterID)
            End If
            '
            Call Main.TestPoint("GetIssueID - IssueID: " & IssueID)
            '
            GetIssueID = IssueID
            '
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetIssueID", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Function GetCurrentIssueID(Main As Object, NewsletterID As Long) As Long
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            '
            CS = Main.OpenCSContent(ContentNameNewsletterIssues, "(PublishDate<=" & Main.EncodeSQLDate(Now()) & ") AND (NewsletterID=" & NewsletterID & ")", "PublishDate desc, ID desc", , "ID")
            If Main.CSOK(CS) Then
                GetCurrentIssueID = Main.GetCSInteger(CS, "ID")
            End If
            Call Main.CloseCS(CS)
            '
            Exit Function
ErrorTrap:
            Call HandleError("NewsLetter", "GetCurrentIssueID", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Function GetUnpublishedIssueList(Main As Object, NewsletterID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            Dim ID As Long
            Dim Name As String
            Dim Active As Boolean
            Dim PublishDate As Date
            Dim Copy As String
            Dim DateAdded As Date
            '
            CS = Main.OpenCSContent(ContentNameNewsletterIssues, "(newsletterid=" & NewsletterID & ")and(PublishDate is null)or(PublishDate>" & Main.EncodeSQLDate(Now()) & ")", "PublishDate desc, ID desc", , "ID")
            Do While Main.CSOK(CS)
                ID = Main.GetCSInteger(CS, "ID")
                Name = Trim(Main.GetCSText(CS, "name"))
                Active = Main.GetCSBoolean(CS, "active")
                PublishDate = Main.GetCSDate(CS, "PublishDate")
                DateAdded = Main.GetCSDate(CS, "DateAdded")
                Copy = Name
                If Copy = "" Then
                    Copy = "unnamed #" & ID
                End If
                If Not Active Then
                    Copy = Copy & ",inactive"
                End If
                If DateAdded <> CDate(0) Then
                    Copy = Copy & ", created " & Int(DateAdded)
                End If
                If PublishDate <> CDate(0) Then
                    Copy = Copy & ", publish " & Int(PublishDate)
                End If
                If Main.IsContentManager("Newsletters") Then
                    Copy = "<a href=""?" & Main.RefreshQueryString & "&" & RequestNameIssueID & "=" & ID & """>" & Copy & "</a>"
                End If
                GetUnpublishedIssueList = GetUnpublishedIssueList & "<li>" & Copy & "</li>"
                Call Main.NextCSRecord(CS)
            Loop
            Call Main.CloseCS(CS)
            '
            If GetUnpublishedIssueList <> "" Then
                GetUnpublishedIssueList = "<UL>" & GetUnpublishedIssueList & "</UL>"
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetUnpublishedIssueList", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Sub UpgradeAddOn(Main As Object)
            On Error GoTo ErrorTrap
            '
            Dim CurrentVersion As String
            Dim AddOnVersion As String
            Dim CSPointer As Long
            Dim NewsletterID As Long
            '
            AddOnVersion = App.Major & "." & App.Minor & "." & App.Revision
            '
            CurrentVersion = Main.GetSiteProperty(PropertyAddOnVersion, "0.0.0")
            '
            If CurrentVersion <> AddOnVersion Then
                '
                If CurrentVersion < "1.0.54" Then
                    '
                    ' Get Default NewsletterID
                    '
                    CSPointer = Main.OpenCSContent(ContentNameNewsletters, "Name=" & Main.EncodeSQLText(DefaultRecord))
                    If Not Main.CSOK(CSPointer) Then
                        Call Main.CloseCS(CSPointer)
                        CSPointer = Main.InsertCSContent(ContentNameNewsletters)
                    End If
                    If Main.CSOK(CSPointer) Then
                        NewsletterID = Main.GetCSInteger(CSPointer, "ID")
                    End If
                    Call Main.CloseCS(CSPointer)
                    '
                    CSPointer = Main.OpenCSContent(ContentNameNewsletterIssues, "NewsletterID is Null")
                    Do While Main.CSOK(CSPointer)
                        Call Main.SetCS(CSPointer, "NewsletterID", NewsletterID)
                        Call Main.NextCSRecord(CSPointer)
                    Loop
                    Call Main.CloseCS(CSPointer)
                End If
                '
                If CurrentVersion < "1.0.93" Then
                    Call Main.ExecuteSQL("Default", "Update NewsletterissuePages Set AllowReadMore=1")
                End If
                '
                Call Main.SetSiteProperty(PropertyAddOnVersion, AddOnVersion)
                '
            End If
            '
            Exit Sub
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "Upgrade", Err.Number, Err.Source, Err.Description, True, False)
        End Sub
        '
        Friend Function GetNewsletterID(Main As Object, NewsletterName As String) As String
            On Error GoTo ErrorTrap
            '
            Dim NewsletterID As Long
            Dim TemplateCopy As String
            Dim TemplateID As Long
            Dim CS As Long
            Dim CSIssue As Long
            Dim AOPointer As Long
            Dim StyleString As String
            '
            CS = Main.OpenCSContent(ContentNameNewsletters, "Name=" & Main.EncodeSQLText(NewsletterName))
            If Main.CSOK(CS) Then
                NewsletterID = Main.GetCSInteger(CS, "ID")
            Else
                Call Main.CloseCS(CS)
                '
                ' moved the entire build newsletter process here - eliminating the optional build step, makes it easier for cm
                ' Build Default Template
                '
                TemplateID = GetDefaultTemplateID(Main)
                CS = Main.OpenCSContent("Newsletter Templates", "id=" & TemplateID)
                If Main.IsCSOK(CS) Then
                    TemplateCopy = Trim(Main.GetCSText(CS, "Template"))
                    If TemplateCopy = "" Then
                        TemplateCopy = GetDefaultTemplateCopy()
                        Call Main.SetCS(CS, "Template", TemplateCopy)
                    End If
                End If
                Call Main.CloseCS(CS)
                '
                ' Build Newsletter
                '
                CS = Main.InsertCSRecord(ContentNameNewsletters)
                If Main.CSOK(CS) Then
                    NewsletterID = Main.GetCSInteger(CS, "ID")
                    Call Main.SetCS(CS, "Name", NewsletterName)
                    Call Main.SetCS(CS, "TemplateID", TemplateID)
                    AOPointer = Main.OpenCSContent("Add-Ons", "ccGUID=" & Main.EncodeSQLText(NewsletterAddonGuid), , , , , "StylesFileName")
                    If Main.CSOK(AOPointer) Then
                        StyleString = Main.GetCS(AOPointer, "StylesFilename")
                    End If
                    Call Main.CloseCS(AOPointer)
                    Call Main.SetCS(CS, "StylesFileName", StyleString)
                End If
                '
                ' Build the first issue in the newsletter
                '
                CSIssue = Main.InsertCSRecord("Newsletter Issues")
                If Main.IsCSOK(CSIssue) Then
                    Call Main.SetCS(CSIssue, "name", NewsletterName & " - Issue 1")
                    Call Main.SetCS(CSIssue, "NewsletterID", NewsletterID)
                    Call Main.SetCS(CSIssue, "PublishDate", Now())
                End If
                Call Main.CloseCS(CSIssue)
            End If
            Call Main.CloseCS(CS)
            '
            GetNewsletterID = NewsletterID
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetNewsletterID", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Sub SortCategoriesByIssue(Main As Object, IssueID As Long)
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            Dim CategoryID As Long
            Dim Sort As Long
            Dim SortUp As Long
            Dim SortDown As Long
            Dim SQL As String
            Dim Pointer As Long
            Dim MainSQL As String
            Dim PreviousID As Long
            Dim PreviousCategoryID As Long
            Dim SortArray As Object
            Dim SortArrayPointer As Long
            Dim SortArrayCount As Long
            Dim SortOrder As String
            Dim RuleCategoryID As Long
            Dim RuleIssueID As Long
            '
            CategoryID = Main.GetStreamInteger(RequestNameSortUp)
            '
            '   Check for Categories without rules, since rules decide sort order of categories, no stories show if
            '       associated to a category without a rule, join fails.
            '
            SQL = "SELECT NIP.CategoryID AS CatID, NewsletterID AS IssueID "
            SQL = SQL & "FROM NewsletterIssuePages NIP "
            SQL = SQL & "WHERE (NIP.CategoryID Not IN (SELECT CategoryID FROM NewsletterIssueCategoryRules WHERE NewsletterIssueID=" & Main.EncodeSQLNumber(IssueID) & ")) "
            SQL = SQL & "AND (NIP.CategoryID Is Not Null)"
            ' 1/19/2009 just look for IssuePages within this issue that do not have IssueCategoryRules for this issue
            SQL = SQL & "AND (NIP.NewsletterID=" & Main.EncodeSQLNumber(IssueID) & ")"
            '
            CS = Main.OpenCSSQL("Default", SQL)
            Do While Main.CSOK(CS)
                Pointer = Main.InsertCSRecord(ContentNameIssueRules)
                If Main.CSOK(Pointer) Then
                    RuleCategoryID = Main.GetCSInteger(CS, "CatID")
                    RuleIssueID = Main.GetCSInteger(CS, "IssueID")
                    SortOrder = GetSortOrder(Main, RuleCategoryID, RuleIssueID)
                    Call Main.SetCS(Pointer, "NewsletterIssueID", RuleIssueID)
                    Call Main.SetCS(Pointer, "Active", 1)
                    Call Main.SetCS(Pointer, "CategoryID", RuleCategoryID)
                    Call Main.SetCS(Pointer, "SortOrder", SortOrder)
                End If
                Call Main.CloseCS(Pointer)
                Call Main.NextCSRecord(CS)
            Loop
            Call Main.CloseCS(CS)
            '
            If CategoryID <> 0 Then
                '
                MainSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder"
                MainSQL = MainSQL & " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR"
                MainSQL = MainSQL & " Where (NIC.ID = NIR.CategoryID)"
                MainSQL = MainSQL & " AND (NIR.NewsletterIssueID=" & IssueID & ")"
                MainSQL = MainSQL & " AND (NIC.Active<>0)"
                MainSQL = MainSQL & " AND (NIR.Active<>0)"
                MainSQL = MainSQL & " ORDER BY NIR.SortOrder"
                '
                CS = Main.OpenCSSQL("Default", MainSQL)
                SortArray = Main.GetCSRows(CS)
                SortArrayCount = UBound(SortArray, 2)
                For SortArrayPointer = 0 To SortArrayCount
                    If (CategoryID = SortArray(0, SortArrayPointer)) And (SortArrayPointer <> 0) Then
                        SortArray(1, SortArrayPointer - 1) = PadValue(Sort, 4)
                        SortArray(1, SortArrayPointer) = PadValue(Sort - 10, 4)
                    Else
                        SortArray(1, SortArrayPointer) = PadValue(Sort, 4)
                    End If
                    Sort = Sort + 10
                Next
                '
                SortArrayPointer = 0
                '
                For SortArrayPointer = 0 To SortArrayCount
                    SQL = "Update NewsletterIssueCategoryRules SET SortOrder=" & SortArray(1, SortArrayPointer) & " WHERE (CategoryID=" & Main.EncodeSQLNumber(SortArray(0, SortArrayPointer)) & ") AND (NewsletterIssueID=" & Main.EncodeSQLNumber(IssueID) & ")"
                    'Call Main.WriteStream("SQL " & SortArrayPointer & ": " & SQL)
                    Call Main.ExecuteSQL("Default", SQL)
                Next
                '
            End If
            '
            Exit Sub
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "SortCategoriesByIssue", Err.Number, Err.Source, Err.Description, True, False)
        End Sub
        '
Friend Function HasArticleAccess(Main As Object, ArticleID As Long, Optional GivenGroupID As Long) As Boolean
            On Error GoTo ErrorTrap
            '
            Dim CSPointer As Long
            Dim AccessFlag As Boolean
            Dim ThisTest As String
            '
            If GivenGroupID <> 0 Then
                CSPointer = Main.OpenCSContent(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                If Not Main.CSOK(CSPointer) Then
                    HasArticleAccess = True
                Else
                    Do While Main.CSOK(CSPointer)
                        If Main.GetCSInteger(CSPointer, "GroupID") = GivenGroupID Then
                            HasArticleAccess = True
                        End If
                        Call Main.NextCSRecord(CSPointer)
                    Loop
                End If
                Call Main.CloseCS(CSPointer)
            Else
                If Not Main.IsContentManager("Newsletters") Then
                    CSPointer = Main.OpenCSContent(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
                    If Not Main.CSOK(CSPointer) Then
                        HasArticleAccess = True
                    Else
                        Do While Main.CSOK(CSPointer)
                            ThisTest = Main.GetCSLookup(CSPointer, "GroupID")
                            '
                            If ThisTest <> "" Then
                                If Main.IsGroupMember(ThisTest) Then
                                    HasArticleAccess = True
                                End If
                            End If
                            Call Main.NextCSRecord(CSPointer)
                        Loop
                    End If
                    Call Main.CloseCS(CSPointer)
                Else
                    HasArticleAccess = True
                End If
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "HasArticleAccess", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Function GetCategoryAccessString(Main As Object, CategoryID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            Dim SQL As String
            Dim Stream As String
            '
            SQL = "SELECT ID "
            SQL = SQL & "From NewsletterIssuePages "
            SQL = SQL & "WHERE (CategoryID=" & Main.EncodeSQLNumber(CategoryID) & ") "
            SQL = SQL & "AND (ID not in(Select NewsletterPageID FROM NewsletterPageGroupRules))"
            '
            ' first scheck for any unblocked story
            '
            CS = Main.OpenCSSQL("Default", SQL)
            If Main.CSOK(CS) Then
                '
                '   no unblocked stories, look for blocked stories
                '
                Call Main.CloseCS(CS)
                SQL = "SELECT GR.GroupID "
                SQL = SQL & "FROM NewsletterPageGroupRules GR, NewsletterIssuePages NIP "
                SQL = SQL & "Where (GR.NewsletterPageID = NIP.ID) "
                SQL = SQL & "AND (NIP.CategoryID=" & Main.EncodeSQLNumber(CategoryID) & ") "
                '
                CS = Main.OpenCSSQL("Default", SQL)
                Do While Main.CSOK(CS)
                    If Stream <> "" Then
                        Stream = Stream & ","
                    End If
                    Stream = Stream & Main.GetCSInteger(CS, "GroupID")
                    Call Main.NextCSRecord(CS)
                Loop
                Call Main.CloseCS(CS)
            End If
            Call Main.CloseCS(CS)
            '
            '    If Stream <> "" Then
            '        Stream = Stream & ","
            '    End If
            '
            GetCategoryAccessString = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetCategoryAccessString", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Function GetArticleAccessString(Main As Object, StoryID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            Dim SQL As String
            Dim Stream As String
            '
            SQL = "SELECT GR.GroupID "
            SQL = SQL & "FROM NewsletterPageGroupRules GR "
            SQL = SQL & "Where (GR.NewsletterPageID=" & Main.EncodeSQLNumber(StoryID) & ")"
            '
            CS = Main.OpenCSSQL("Default", SQL)
            Do While Main.CSOK(CS)
                If Stream <> "" Then
                    Stream = Stream & ","
                End If
                Stream = Stream & Main.GetCSInteger(CS, "GroupID")
                Call Main.NextCSRecord(CS)
            Loop
            Call Main.CloseCS(CS)
            '
            '    If Stream <> "" Then
            '        Stream = Stream & ","
            '    End If
            '
            GetArticleAccessString = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetArticleAccessString", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Friend Function HasAccess(Main As Object, GroupString As String) As Boolean
            On Error GoTo ErrorTrap
            '
            Dim ListArray() As String
            Dim ListArrayCount As Long
            Dim ListArrayPointer As Long
            Dim AccessFlag As Boolean
            '
            If Main.IsContentManager("Newsletters") Then
                HasAccess = True
            Else
                If GroupString <> "" Then
                    If InStr(1, GroupString, ",", vbTextCompare) <> 0 Then
                        ListArray() = Split(GroupString, ",", , vbTextCompare)
                        ListArrayCount = UBound(ListArray())
                        For ListArrayPointer = 0 To ListArrayCount
                            If Main.IsGroupMember(Main.GetRecordName("Groups", ListArray(ListArrayPointer))) Then
                                HasAccess = True
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    HasAccess = True
                End If
            End If
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetArticleAccessString", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function PadValue(Value As Long, StringLenghth As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim Counter As Long
            Dim ValueLenghth As Long
            Dim InnerValue As String
            Dim Stream As String
            '
            InnerValue = CStr(Value)
            ValueLenghth = Len(InnerValue)
            '
            If ValueLenghth < StringLenghth Then
                For Counter = ValueLenghth To StringLenghth - 1
                    InnerValue = "0" & InnerValue
                Next
            End If
            '
            PadValue = InnerValue
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "PadValue", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        Private Function GetSortOrder(Main As Object, CategoryID As Long, IssueID As Long) As String
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            Dim Stream As String
            '
            CS = Main.OpenCSContent("Newsletter Issue Category Rules", "(CategoryID=" & CategoryID & ") AND (NewsletterIssueID=" & IssueID & ")", , , , , "SortOrder")
            If Main.CSOK(CS) Then
                Stream = Main.GetCSText(CS, "SortOrder")
            End If
            Call Main.CloseCS(CS)
            '
            If Stream = "" Then
                Stream = "0"
            End If
            '
            GetSortOrder = Stream
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetSortOrder", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        '
        '
        Friend Function GetDefaultTemplateID(Main As Object) As Long
            On Error GoTo ErrorTrap
            '
            Dim CS As Long
            Dim TemplateID As Long
            Dim TemplateCopy As String
            '
            ' try default template
            '
            CS = Main.OpenCSContent("Newsletter Templates", "name=" & KmaEncodeSQLText("Default"))
            If Main.IsCSOK(CS) Then
                '
                ' Use the default template in their Db already
                '
                TemplateID = Main.GetCSInteger(CS, "ID")
                TemplateCopy = Trim(Main.GetCSText(CS, "Template"))
                If TemplateCopy = "" Then
                    TemplateCopy = DefaultTemplate
                    TemplateCopy = Replace(TemplateCopy, "{{ACID0}}", GetRandomInteger())
                    TemplateCopy = Replace(TemplateCopy, "{{ACID1}}", GetRandomInteger())
                    Call Main.SetCS(CS, "Template", TemplateCopy)
                End If
            End If
            Call Main.CloseCS(CS)
            If TemplateID = 0 Then
                '
                ' build default template
                '
                TemplateCopy = GetDefaultTemplateCopy
                CS = Main.InsertCSRecord("Newsletter Templates")
                If Main.IsCSOK(CS) Then
                    Call Main.SetCS(CS, "name", "Default")
                    Call Main.SetCS(CS, "Template", TemplateCopy)
                    TemplateID = Main.GetCSInteger(CS, "ID")
                End If
                Call Main.CloseCS(CS)
            End If
            '
            GetDefaultTemplateID = TemplateID
            '
            Exit Function
ErrorTrap:
            Call HandleError("aoNewsletter.CommonClass", "GetDefaultTemplateID", Err.Number, Err.Source, Err.Description, True, False)
        End Function
        '
        '
        '
        Friend Function GetDefaultTemplateCopy() As String
            GetDefaultTemplateCopy = DefaultTemplate
            GetDefaultTemplateCopy = Replace(GetDefaultTemplateCopy, "{{ACID0}}", GetRandomInteger())
            GetDefaultTemplateCopy = Replace(GetDefaultTemplateCopy, "{{ACID1}}", GetRandomInteger())
        End Function
        ''
        ''
        ''
        'Friend Function GetAuthoringLinks(Main As Object, IssuePageID As Long, IssueID As Long, NewsletterID As Long, WorkingQueryStringPlus) As String
        '    On Error GoTo ErrorTrap
        '    '
        '    Dim CSPointer As Long
        '    Dim QS As String
        '    '
        '    If Main.IsLinkAuthoring(ContentNameNewsletterIssues) Then
        '        ' 1/1/09 added admin hint wrapper
        '        'GetAuthoringLinks = "<br /><br />"
        '        If IssuePageID <> 0 Then
        '            GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssuePages) & "&af=4&id=" & IssuePageID & "&" & ReferLink & """>Edit this page</a></div>"
        '        End If
        '        If IssueID <> 0 Then
        '            GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssuePages) & "&af=4&aa=2&ad=1&wc=NewsletterID=" & IssueID & "&" & ReferLink & """>Add a new page</a></div>"
        '            GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssues) & "&af=4&id=" & IssueID & "&" & ReferLink & """>Edit this issue</a></div>"
        '        End If
        '        GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameNewsletterIssues) & "&wl0=newsletterid&wr0=" & NewsletterID & "&af=4&aa=2&ad=1&" & "&" & ReferLink & """>Add a new issue</a></div>"
        '        If IssueID <> 0 Then
        '            GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink""><a href=""" & WorkingQueryStringPlus & RequestNameFormID & "=" & FormEmail & "&" & RequestNameIssueID & "=" & IssueID & """>Create&nbsp;email&nbsp;version</a></div>"
        '        End If
        '        GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink""><a href = ""http://" & Main.ServerHost & Main.SiteProperty_AdminURL & "?cid=" & Main.GetContentID(ContentNameIssueCategories) & "&" & ReferLink & """>Edit categories</a></div>"
        '        '
        '        ' Future Issues
        '        '
        '        CSPointer = Main.OpenCSContent(ContentNameNewsletterIssues, "(PublishDate>" & Main.EncodeSQLDate(Now()) & ") OR (PublishDate is Null) OR (PublishDate=" & KmaEncodeSQLDate(0) & ")", "PublishDate desc")
        '        If Main.CSOK(CSPointer) Then
        '            GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink"">Unpublished Issues</div>"
        '            Do While Main.CSOK(CSPointer)
        '                ' 1/1/09 - JK - always linked to root path, also added qs incase request name was already in wqsp
        '                QS = WorkingQueryStringPlus
        '                QS = ModifyQueryString(QS, RequestNameIssueID, Main.GetCSInteger(CSPointer, "ID"), True)
        '                GetAuthoringLinks = GetAuthoringLinks & "<div class=""AdminLink"" style=""padding-left:5px""><a href=""" & QS & """>" & Main.GetCSText(CSPointer, "Name") & "</a></div>"
        '                'Stream = Stream & "<div class=""PageList""><a href=""http://" & Main.ServerHost & Main.ServerAppRootPath & Main.ServerPage & WorkingQueryStringPlus & RequestNameIssueID & "=" & Main.GetCSInteger(CSPointer, "ID") & """>" & Main.GetCSText(CSPointer, "Name") & "</a></div>"
        '                Call Main.NextCSRecord(CSPointer)
        '            Loop
        '        End If
        '        Call Main.CloseCS(CSPointer)
        '    End If
        '
        '    '
        '    Exit Function
        'ErrorTrap:
        '    Call HandleError("aoNewsletter.CommonClass", "GetAuthoringLinks", Err.Number, Err.Source, Err.Description, True, False)
        'End Function

    End Class
End Namespace
