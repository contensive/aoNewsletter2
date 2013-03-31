
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace newsletter2
    '
    ' Sample Vb addon
    '
    Public Class newsletterCommonClass
        '
        Public Const cr As String = vbCrLf & vbTab
        '
        Private Private_LegacySiteSites_Loaded As Boolean
        Private Private_LegacySiteSites As String
        Private EditWrapperCnt As Integer = 0
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub handleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterCommonClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        '
        '
        Friend Function GetIssueID(cp As CPBaseClass, NewsletterID As Integer) As Integer
            'On Error GoTo ErrorTrap
            '
            Dim IssueID As Integer
            '
            IssueID = cp.Doc.GetInteger(RequestNameIssueID)
            '
            Call cp.Site.TestPoint("GetIssueID - IssueID From Stream: " & IssueID)
            Call cp.Site.TestPoint("GetIssueID - NewsletterID: " & NewsletterID)
            '
            If IssueID = 0 Then
                IssueID = GetCurrentIssueID(cp, NewsletterID)
            End If
            '
            Call cp.Site.TestPoint("GetIssueID - IssueID: " & IssueID)
            '
            GetIssueID = IssueID
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetIssueID")
        End Function
        '
        Friend Function GetCurrentIssueID(cp As CPBaseClass, NewsletterID As Integer) As Integer
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            '
            Call cs.Open(ContentNameNewsletterIssues, "(PublishDate<=" & cp.Db.EncodeSQLDate(Now()) & ") AND (NewsletterID=" & NewsletterID & ")", "PublishDate desc, ID desc", , "ID")
            If cs.OK() Then
                GetCurrentIssueID = cs.GetInteger("ID")
            End If
            Call cs.Close()
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError(cp, ex, "GetCurrentIssueID")
        End Function
        '
        Friend Function GetUnpublishedIssueList(cp As CPBaseClass, NewsletterID As Integer, cn As newsletterCommonClass) As String
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim ID As Integer
            Dim Name As String
            Dim Active As Boolean
            Dim PublishDate As Date
            Dim Copy As String
            Dim DateAdded As Date
            '
            Call cs.Open(ContentNameNewsletterIssues, "(newsletterid=" & NewsletterID & ")and(PublishDate is null)or(PublishDate>" & cp.Db.EncodeSQLDate(Now()) & ")", "PublishDate desc, ID desc", , "ID")
            Do While cs.OK()
                ID = cs.GetInteger("ID")
                Name = Trim(cs.GetText("name"))
                Active = cs.GetBoolean("active")
                PublishDate = cs.GetDate("PublishDate")
                DateAdded = cs.GetDate("DateAdded")
                Copy = Name
                If Copy = "" Then
                    Copy = "unnamed #" & ID
                End If
                If Not Active Then
                    Copy = Copy & ",inactive"
                End If
                If cn.encodeMinDate(DateAdded) <> Date.MinValue Then
                    Copy = Copy & ", created " & Int(DateAdded)
                End If
                If PublishDate <> Date.MinValue Then
                    Copy = Copy & ", publish " & Int(PublishDate)
                End If
                If cp.User.IsContentManager("Newsletters") Then
                    Copy = "<a href=""?" & cp.Doc.RefreshQueryString & "&" & RequestNameIssueID & "=" & ID & """>" & Copy & "</a>"
                End If
                GetUnpublishedIssueList = GetUnpublishedIssueList & "<li>" & Copy & "</li>"
                Call cs.GoNext()
            Loop
            Call cs.Close()
            '
            If GetUnpublishedIssueList <> "" Then
                GetUnpublishedIssueList = "<UL>" & GetUnpublishedIssueList & "</UL>"
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("aoNewsletter.newsletterCommonClass", "GetUnpublishedIssueList")
        End Function
        '
        Friend Function GetNewsletterID(cp As CPBaseClass, NewsletterName As String) As String
            'On Error GoTo ErrorTrap
            '
            Dim NewsletterID As Integer
            Dim TemplateCopy As String
            Dim TemplateID As Integer
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim CSIssue As CPCSBaseClass = cp.CSNew()
            Dim AOPointer As CPCSBaseClass = cp.CSNew()
            Dim StyleString As String
            '
            Call cs.Open(ContentNameNewsletters, "Name=" & cp.Db.EncodeSQLText(NewsletterName))
            If cs.OK() Then
                NewsletterID = cs.GetInteger("ID")
            Else
                Call cs.Close()
                '
                ' moved the entire build newsletter process here - eliminating the optional build step, makes it easier for cm
                ' Build Default Template
                '
                TemplateID = GetDefaultTemplateID(cp)
                Call cs.Open("Newsletter Templates", "id=" & TemplateID)
                If cs.OK() Then
                    TemplateCopy = Trim(cs.GetText("Template"))
                    If TemplateCopy = "" Then
                        TemplateCopy = GetDefaultTemplateCopy(cp)
                        Call cs.SetField("Template", TemplateCopy)
                    End If
                End If
                Call cs.Close()
                '
                ' Build Newsletter
                '
                Call cs.Insert(ContentNameNewsletters)
                If cs.OK() Then
                    NewsletterID = cs.GetInteger("ID")
                    Call cs.SetField("Name", NewsletterName)
                    Call cs.SetField("TemplateID", TemplateID)
                    Call AOPointer.Open("Add-Ons", "ccGUID=" & cp.Db.EncodeSQLText(NewsletterAddonGuid), , , , , "StylesFileName")
                    If AOPointer.OK Then
                        StyleString = AOPointer.GetText("StylesFilename")
                    End If
                    Call AOPointer.Close()
                    Call cs.SetField("StylesFileName", StyleString)
                End If
                '
                ' Build the first issue in the newsletter
                '
                Call CSIssue.Insert("Newsletter Issues")
                If CSIssue.OK() Then
                    Call CSIssue.SetField("name", NewsletterName & " - Issue 1")
                    Call CSIssue.SetField("NewsletterID", NewsletterID)
                    Call CSIssue.SetField("PublishDate", Now())
                End If
                Call CSIssue.Close()
            End If
            Call cs.Close()
            '
            GetNewsletterID = NewsletterID
        End Function
        '
        Friend Sub SortCategoriesByIssue(cp As CPBaseClass, IssueID As Integer)
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim Pointer As CPCSBaseClass = cp.CSNew()
            Dim CategoryID As Integer
            Dim Sort As Integer
            Dim SortUp As Integer
            Dim SortDown As Integer
            Dim SQL As String
            Dim MainSQL As String
            Dim PreviousID As Integer
            Dim PreviousCategoryID As Integer
            Dim SortArray As Object
            Dim SortArrayPointer As Integer
            Dim SortArrayCount As Integer
            Dim SortOrder As String
            Dim RuleCategoryID As Integer
            Dim RuleIssueID As Integer
            '
            CategoryID = cp.Doc.GetInteger(RequestNameSortUp)
            '
            '   Check for Categories without rules, since rules decide sort order of categories, no stories show if
            '       associated to a category without a rule, join fails.
            '
            SQL = "SELECT NIP.CategoryID AS CatID, NewsletterID AS IssueID "
            SQL = SQL & "FROM NewsletterIssuePages NIP "
            SQL = SQL & "WHERE (NIP.CategoryID Not IN (SELECT CategoryID FROM NewsletterIssueCategoryRules WHERE NewsletterIssueID=" & cp.Db.EncodeSQLNumber(IssueID) & ")) "
            SQL = SQL & "AND (NIP.CategoryID Is Not Null)"
            ' 1/19/2009 just look for IssuePages within this issue that do not have IssueCategoryRules for this issue
            SQL = SQL & "AND (NIP.NewsletterID=" & cp.Db.EncodeSQLNumber(IssueID) & ")"
            '
            Call cs.OpenSQL(SQL)
            Do While cs.OK()
                Call Pointer.Insert(ContentNameIssueRules)
                If Pointer.OK Then
                    RuleCategoryID = cs.GetInteger("CatID")
                    RuleIssueID = cs.GetInteger("IssueID")
                    SortOrder = GetSortOrder(cp, RuleCategoryID, RuleIssueID)
                    Call Pointer.SetField("NewsletterIssueID", RuleIssueID)
                    Call Pointer.SetField("Active", 1)
                    Call Pointer.SetField("CategoryID", RuleCategoryID)
                    Call Pointer.SetField("SortOrder", SortOrder)
                End If
                Call Pointer.GoNext()
                Call cs.GoNext()
            Loop
            Call cs.Close()
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
                ' b/c cp has no cp.getRows
                '
                If cs.OpenSQL(MainSQL) Then
                    SortArrayCount = cs.GetRowCount
                    ReDim SortArray(2, SortArrayCount)
                    Dim ptr As Integer = 0
                    Do While cs.OK()
                        SortArray(0, ptr) = cs.GetText("categoryId")
                        SortArray(1, ptr) = cs.GetText("sortOrder")
                        Call cs.GoNext()
                    Loop
                End If
                'SortArray = cs.Main.GetCSRows(cs)
                '
                SortArrayCount = UBound(SortArray, 2)
                For SortArrayPointer = 0 To SortArrayCount
                    If (CategoryID = SortArray(0, SortArrayPointer)) And (SortArrayPointer <> 0) Then
                        SortArray(1, SortArrayPointer - 1) = PadValue(cp, Sort, 4)
                        SortArray(1, SortArrayPointer) = PadValue(cp, Sort - 10, 4)
                    Else
                        SortArray(1, SortArrayPointer) = PadValue(cp, Sort, 4)
                    End If
                    Sort = Sort + 10
                Next
                '
                SortArrayPointer = 0
                '
                For SortArrayPointer = 0 To SortArrayCount
                    SQL = "Update NewsletterIssueCategoryRules SET SortOrder=" & SortArray(1, SortArrayPointer) & " WHERE (CategoryID=" & cp.Db.EncodeSQLNumber(SortArray(0, SortArrayPointer)) & ") AND (NewsletterIssueID=" & cp.Db.EncodeSQLNumber(IssueID) & ")"
                    'Call Main.WriteStream("SQL " & SortArrayPointer & ": " & SQL)
                    Call cp.Db.ExecuteSQL(SQL)
                Next
                '
            End If
        End Sub
        '
        Friend Function HasArticleAccess(cp As CPBaseClass, ArticleID As Integer, Optional GivenGroupID As Integer = 0) As Boolean
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim ThisTest As String
            '
            If GivenGroupID <> 0 Then
                Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , "GroupID")
                If Not cs.OK Then
                    HasArticleAccess = True
                Else
                    Do While cs.OK
                        If cs.GetInteger("GroupID") = GivenGroupID Then
                            HasArticleAccess = True
                        End If
                        Call cs.GoNext()
                    Loop
                End If
                Call cs.Close()
            Else
                If Not cp.User.IsContentManager("Newsletters") Then
                    Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , "GroupID")
                    If Not cs.OK Then
                        HasArticleAccess = True
                    Else
                        Do While cs.OK
                            ThisTest = cs.GetText("GroupID")
                            'ThisTest = cs.getText( "GroupID")
                            '
                            If ThisTest <> "" Then
                                If cp.User.IsInGroup(ThisTest) Then
                                    HasArticleAccess = True
                                End If
                            End If
                            Call cs.GoNext()
                        Loop
                    End If
                    Call cs.Close()
                Else
                    HasArticleAccess = True
                End If
            End If
        End Function
        '
        Friend Function GetCategoryAccessString(cp As CPBaseClass, CategoryID As Integer) As String
            'On Error GoTo ErrorTrap
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim SQL As String
            Dim Stream As String
            '
            SQL = "SELECT ID "
            SQL = SQL & "From NewsletterIssuePages "
            SQL = SQL & "WHERE (CategoryID=" & cp.Db.EncodeSQLNumber(CategoryID) & ") "
            SQL = SQL & "AND (ID not in(Select NewsletterPageID FROM NewsletterPageGroupRules))"
            '
            ' first scheck for any unblocked story
            '
            Call cs.OpenSQL(SQL)
            If cs.OK() Then
                '
                '   no unblocked stories, look for blocked stories
                '
                Call cs.Close()
                SQL = "SELECT GR.GroupID "
                SQL = SQL & "FROM NewsletterPageGroupRules GR, NewsletterIssuePages NIP "
                SQL = SQL & "Where (GR.NewsletterPageID = NIP.ID) "
                SQL = SQL & "AND (NIP.CategoryID=" & cp.Db.EncodeSQLNumber(CategoryID) & ") "
                '
                Call cs.OpenSQL(SQL)
                Do While cs.OK()
                    If Stream <> "" Then
                        Stream &= ","
                    End If
                    Stream &= cs.GetInteger("GroupID")
                    Call cs.GoNext()
                Loop
                Call cs.Close()
            End If
            Call cs.Close()
            '
            '    If Stream <> "" Then
            '        stream &=  ","
            '    End If
            '
            GetCategoryAccessString = Stream
        End Function
        '
        Friend Function GetArticleAccessString(cp As CPBaseClass, StoryID As Integer) As String
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim SQL As String
            Dim Stream As String
            '
            SQL = "SELECT GR.GroupID "
            SQL = SQL & "FROM NewsletterPageGroupRules GR "
            SQL = SQL & "Where (GR.NewsletterPageID=" & cp.Db.EncodeSQLNumber(StoryID) & ")"
            '
            Call cs.OpenSQL(SQL)
            Do While cs.OK()
                If Stream <> "" Then
                    Stream &= ","
                End If
                Stream &= cs.GetInteger("GroupID")
                Call cs.GoNext()
            Loop
            Call cs.Close()
            '
            '    If Stream <> "" Then
            '        stream &=  ","
            '    End If
            '
            GetArticleAccessString = Stream
        End Function
        '
        Friend Function HasAccess(cp As CPBaseClass, GroupString As String) As Boolean
            '
            Dim ListArray() As String
            Dim ListArrayCount As Integer
            Dim ListArrayPointer As Integer
            Dim AccessFlag As Boolean
            '
            If cp.User.IsContentManager("Newsletters") Then
                HasAccess = True
            Else
                If GroupString <> "" Then
                    If InStr(1, GroupString, ",", vbTextCompare) <> 0 Then
                        ListArray = Split(GroupString, ",", , vbTextCompare)
                        ListArrayCount = UBound(ListArray)
                        For ListArrayPointer = 0 To ListArrayCount
                            If cp.User.IsInGroup(cp.Content.GetRecordName("Groups", ListArray(ListArrayPointer))) Then
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
            'Exit Function
            'ErrorTrap:
            'Call HandleError("aoNewsletter.newsletterCommonClass", "GetArticleAccessString")
        End Function
        '
        Private Function PadValue(cp As CPBaseClass, Value As Integer, StringLenghth As Integer) As String
            Dim Counter As Integer
            Dim ValueLenghth As Integer
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
        End Function
        '
        Private Function GetSortOrder(cp As CPBaseClass, CategoryID As Integer, IssueID As Integer) As String
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim Stream As String
            '
            Call cs.Open("Newsletter Issue Category Rules", "(CategoryID=" & CategoryID & ") AND (NewsletterIssueID=" & IssueID & ")", , , , , "SortOrder")
            If cs.OK() Then
                Stream = cs.GetText("SortOrder")
            End If
            Call cs.Close()
            '
            If Stream = "" Then
                Stream = "0"
            End If
            '
            GetSortOrder = Stream
        End Function
        '
        Friend Function GetDefaultTemplateID(cp As CPBaseClass) As Integer
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim TemplateID As Integer
            Dim TemplateCopy As String
            '
            ' try default template
            '
            Call cs.Open("Newsletter Templates", "name=" & cp.Db.EncodeSQLText("Default"))
            If cs.OK() Then
                '
                ' Use the default template in their Db already
                '
                TemplateID = cs.GetInteger("ID")
                TemplateCopy = Trim(cs.GetText("Template"))
                If TemplateCopy = "" Then
                    TemplateCopy = DefaultTemplate
                    TemplateCopy = Replace(TemplateCopy, "{{ACID0}}", GetRandomInteger())
                    TemplateCopy = Replace(TemplateCopy, "{{ACID1}}", GetRandomInteger())
                    Call cs.SetField("Template", TemplateCopy)
                End If
            End If
            Call cs.Close()
            If TemplateID = 0 Then
                '
                ' build default template
                '
                TemplateCopy = GetDefaultTemplateCopy(cp)
                Call cs.Insert("Newsletter Templates")
                If cs.OK() Then
                    Call cs.SetField("name", "Default")
                    Call cs.SetField("Template", TemplateCopy)
                    TemplateID = cs.GetInteger("ID")
                End If
                Call cs.Close()
            End If
            '
            GetDefaultTemplateID = TemplateID
        End Function
        '
        Friend Function GetDefaultTemplateCopy(cp As CPBaseClass) As String
            GetDefaultTemplateCopy = DefaultTemplate
            GetDefaultTemplateCopy = Replace(GetDefaultTemplateCopy, "{{ACID0}}", GetRandomInteger())
            GetDefaultTemplateCopy = Replace(GetDefaultTemplateCopy, "{{ACID1}}", GetRandomInteger())
        End Function
        '
        '===================================================================================================
        '   Wrap the content in a common wrapper if authoring is enabled
        '===================================================================================================
        '
        Public Function GetEditWrapper(cp As CPBaseClass, Caption As String, Content As String) As String
            Dim returnString As String = Content
            Try
                '
                Dim IsAuthoring As Boolean
                '
                IsAuthoring = cp.User.IsEditingAnything()
                If Not IsAuthoring Then
                    returnString = Content
                Else
                    returnString = GetLegacySiteStyles()
                    returnString = returnString _
                    & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccEditWrapper"">"
                    If Caption <> "" Then
                        returnString = returnString _
                                & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccEditWrapperCaption"">" _
                                & Caption _
                                & "<!-- <img alt=""space"" src=""/ccLib/images/spacer.gif"" width=1 height=22 align=absmiddle> -->" _
                                & "</td></tr></table>"
                    End If
                    returnString = returnString _
                            & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccEditWrapperContent"" id=""editWrapper" & EditWrapperCnt & """>" _
                            & Content _
                            & "</td></tr></table>" _
                        & "</td></tr></table>"
                    EditWrapperCnt = EditWrapperCnt + 1
                End If
            Catch ex As Exception

            End Try
            Return returnString
        End Function
        '
        '===================================================================================================
        '   Wrap the content in a common wrapper if authoring is enabled
        '===================================================================================================
        '
        Public Function GetAdminHintWrapper(cp As CPBaseClass, Content As String) As String
            Dim returnString As String = Content
            Try
                If cp.User.IsEditingAnything() Or cp.User.IsAdmin Then
                    returnString = GetLegacySiteStyles() _
                        & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccHintWrapper"">" _
                            & "<table border=0 width=""100%"" cellspacing=0 cellpadding=0><tr><td class=""ccHintWrapperContent"">" _
                            & "<b>Administrator</b>" _
                            & "<BR>" _
                            & "<BR>" & Content _
                            & "</td></tr></table>" _
                        & "</td></tr></table>"
                End If
            Catch ex As Exception

            End Try
            Return returnString
        End Function
        '
        Friend Function encodeMinDate(source As Date) As Date
            Dim returnDate As Date = source
            If returnDate < CDate("1/1/1990") Then
                returnDate = Date.MinValue
            End If
        End Function
        '
        '=================================================================================
        '   Get a Random Long Value
        '=================================================================================
        '
        Public Function GetRandomInteger() As Long
            Dim RandomLimit As Long
            RandomLimit = 32767
            Randomize()
            GetRandomInteger = (Rnd() * RandomLimit)
        End Function
        '
        '
        '
        Private Function GetLegacySiteStyles()
            If Not Private_LegacySiteSites_Loaded Then
                Private_LegacySiteSites_Loaded = True
                '
                ' compatibility with old sites - if they do not get the default style sheet, put it in here
                '
                    GetLegacySiteStyles = "" _
                    & cr & "<!-- compatibility with legacy framework --><style type=text/css>" _
                    & cr & " .ccEditWrapper {border:1px dashed #808080;}" _
                    & cr & " .ccEditWrapperCaption {text-align:left;border-bottom:1px solid #808080;padding:4px;background-color:#40C040;color:black;}" _
                    & cr & " .ccEditWrapperContent{padding:4px;}" _
                    & cr & " .ccHintWrapper {border:1px dashed #808080;margin-bottom:10px}" _
                    & cr & " .ccHintWrapperContent{padding:10px;background-color:#80E080;color:black;}" _
                    & "</style>"
            End If
        End Function
    End Class
End Namespace
