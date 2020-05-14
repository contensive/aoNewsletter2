
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Microsoft.SqlServer.Server

Namespace Controllers
    Public Class NewsletterController
        '
        Public Const cr As String = vbCrLf & vbTab
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Shared Sub handleError(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterCommonClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        Friend Shared Function GetIssueID(ByVal cp As CPBaseClass, ByVal NewsletterID As Integer, ByVal currentIssueId As Integer) As Integer
            '
            Dim IssueID As Integer = cp.Doc.GetInteger(RequestNameIssueID)
            '
            If IssueID = 0 Then
                IssueID = currentIssueId
            End If
            '
            Return IssueID
        End Function
        '
        Friend Shared Function GetCurrentIssueID(cp As CPBaseClass, NewsletterID As Integer) As Integer
            Try
                Dim returnId As Integer = 0
                Using cs As CPCSBaseClass = cp.CSNew()
                    '
                    Call cs.Open(ContentNameNewsletterIssues, "(PublishDate<=" & cp.Db.EncodeSQLDate(Now()) & ") AND (NewsletterID=" & NewsletterID & ")", "PublishDate desc, ID desc")
                    If cs.OK() Then
                        returnId = cs.GetInteger("ID")
                    End If
                    Call cs.Close()
                End Using
                '
                If returnId = 0 Then
                    '
                    ' there are no issues of this newsletter -- create a default issue
                    Return createDefaultIssueGetId(cp, NewsletterID)
                End If
                Return returnId
            Catch ex As Exception
                Call handleError(cp, ex, "getCurrentIssueId")
                Return 0
            End Try
        End Function
        '
        Friend Shared Function GetUnpublishedIssueList(cp As CPBaseClass, NewsletterID As Integer, cn As NewsletterController) As String
            GetUnpublishedIssueList = ""
            '
            Dim qs As String = ""
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim ID As Integer
            Dim Name As String
            Dim Active As Boolean
            Dim PublishDate As Date
            Dim Copy As String
            Dim DateAdded As Date
            Dim isContentMan As Boolean
            '
            isContentMan = cp.User.IsContentManager("Newsletters")
            Call cs.Open(ContentNameNewsletterIssues, "(newsletterid=" & NewsletterID & ")and(PublishDate is null)or(PublishDate>" & cp.Db.EncodeSQLDate(Now()) & ")", "PublishDate desc, ID desc")
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
                If NewsletterController.encodeMinDate(DateAdded) <> Date.MinValue Then
                    Copy = Copy & ", created " & DateAdded.ToShortDateString
                End If
                If PublishDate <> Date.MinValue Then
                    Copy = Copy & ", publish " & PublishDate.ToShortDateString
                End If
                If isContentMan Then
                    qs = cp.Doc.RefreshQueryString
                    qs = cp.Utils.ModifyQueryString(qs, RequestNameIssueID, ID.ToString())
                    Copy = "<a href=""?" & qs & """>" & Copy & "</a>"
                End If
                GetUnpublishedIssueList = GetUnpublishedIssueList & "<li>" & Copy & "</li>"
                Call cs.GoNext()
            Loop
            Call cs.Close()
            '
            If GetUnpublishedIssueList <> "" Then
                If isContentMan Then
                    qs = cp.Doc.RefreshQueryString
                    qs = cp.Utils.ModifyQueryString(qs, RequestNameIssueID, "")
                    GetUnpublishedIssueList &= "<li><a href=""?" & qs & """>Current Issue</a></li>"
                End If
                GetUnpublishedIssueList = "<UL>" & GetUnpublishedIssueList & "</UL>"
            End If
            '
            'Exit Function
            'ErrorTrap:
            'Call HandleError("aoNewsletter.newsletterCommonClass", "GetUnpublishedIssueList")
        End Function
        '
        Friend Shared Function getNewsletterId(cp As CPBaseClass, addonArgInstanceGuid As String) As Integer
            Dim returnId As Integer = 0
            Try
                Dim addonArgumentNewsletterId As Integer = cp.Doc.GetInteger("Newsletter")
                Dim criteria As String = ""
                If addonArgumentNewsletterId > 0 Then
                    criteria = "(id=" & addonArgumentNewsletterId & ")"
                ElseIf (addonArgInstanceGuid <> "") Then
                    criteria = "(ccguid='" & addonArgInstanceGuid & "')"
                Else
                    criteria = "(name='Default')"
                End If
                Using cs As CPCSBaseClass = cp.CSNew()
                    If cs.Open(ContentNameNewsletters, criteria) Then
                        returnId = cs.GetInteger("ID")
                    Else
                        Call cs.Close()
                        '
                        ' must create new newsletter
                        '
                        Call cs.Insert(ContentNameNewsletters)
                        If cs.OK() Then
                            returnId = cs.GetInteger("ID")
                            Dim templateID As Integer = verifyDefaultTemplateGetId(cp)
                            Dim emailTemplateID As Integer = verifyDefaultEmailTemplateGetId(cp)
                            If (addonArgInstanceGuid <> "") Then
                                '
                                ' newsletter called out by guid but not found
                                '
                                Call cs.SetField("ccguid", addonArgInstanceGuid)
                                Call cs.SetField("Name", cp.Content.GetRecordName("page content", cp.Doc.PageId))
                            Else
                                '
                                ' all other cases
                                '
                                Call cs.SetField("Name", "'Default'")
                            End If
                            Call cs.SetField("TemplateID", templateID.ToString())
                            Call cs.SetField("emailTemplateID", emailTemplateID.ToString())
                        End If
                        Call createDefaultIssueGetId(cp, returnId)
                    End If
                    Call cs.Close()
                End Using
            Catch ex As Exception
                handleError(cp, ex, "getnewsletterId")
            End Try
            Return returnId
        End Function
        '
        Friend Shared Function createDefaultIssueGetId(cp As CPBaseClass, newsletterId As Integer) As Integer
            Dim returnId As Integer = 0
            Try
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim StyleString As String = ""
                Dim newsletterName As String = cp.Content.GetRecordName(ContentNameNewsletters, newsletterId)
                '
                ' Build the first issue in the newsletter
                '
                Call cs.Insert("Newsletter Issues")
                If cs.OK() Then
                    returnId = cs.GetInteger("id")
                    Call cs.SetField("name", newsletterName & " Newsletter, Issue 1")
                    Call cs.SetField("NewsletterID", newsletterId.ToString())
                    Call cs.SetField("PublishDate", Now().ToShortDateString)
                    Call cs.SetField("cover", cp.WwwFiles.Read("newsletteraddon\Newsletter Issue Default.html"))
                End If
                Call cs.Close()
                '
                ' Build the first story
                '
                Call cs.Insert("Newsletter Stories")
                If cs.OK() Then
                    Call cs.SetField("name", "The First Story")
                    Call cs.SetField("newsletterid", returnId.ToString())
                    Call cs.SetField("Overview", cp.WwwFiles.Read("newsletteraddon\Newsletter Default Story Overview.html"))
                    Call cs.SetField("body", cp.WwwFiles.Read("newsletteraddon\Newsletter Default Story Copy.html"))
                End If
                Call cs.Close()
            Catch ex As Exception
                Call handleError(cp, ex, "createDefaultIssueGetId")
            End Try
            Return returnId
        End Function
        '
        Friend Shared Sub SortCategoriesByIssue(cp As CPBaseClass, IssueID As Integer)
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim Pointer As CPCSBaseClass = cp.CSNew()
            Dim CategoryID As Integer
            Dim Sort As Integer
            Dim SQL As String
            Dim MainSQL As String
            Dim SortArrayPointer As Integer
            Dim SortArrayCount As Integer
            Dim SortOrder As String
            Dim RuleCategoryID As Integer
            Dim RuleIssueID As Integer
            Dim ptr As Integer = 0
            Dim SortArray(2, 1) As String
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
                    Call Pointer.SetField("NewsletterIssueID", RuleIssueID.ToString())
                    Call Pointer.SetField("Active", "1")
                    Call Pointer.SetField("CategoryID", RuleCategoryID.ToString())
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
                    If SortArrayCount > 0 Then
                        ReDim SortArray(2, SortArrayCount - 1)
                        Do While cs.OK()
                            SortArray(0, ptr) = cs.GetText("categoryId")
                            SortArray(1, ptr) = cs.GetText("sortOrder")
                            ptr += 1
                            Call cs.GoNext()
                        Loop
                        SortArrayCount = ptr
                        For SortArrayPointer = 0 To SortArrayCount - 1
                            If (CategoryID = cp.Utils.EncodeInteger(SortArray(0, SortArrayPointer))) And (SortArrayPointer <> 0) Then
                                SortArray(1, SortArrayPointer - 1) = PadValue(cp, Sort, 4)
                                SortArray(1, SortArrayPointer) = PadValue(cp, Sort - 10, 4)
                            Else
                                SortArray(1, SortArrayPointer) = PadValue(cp, Sort, 4)
                            End If
                            Sort = Sort + 10
                        Next
                        SortArrayPointer = 0
                        For SortArrayPointer = 0 To SortArrayCount - 1
                            SQL = "Update NewsletterIssueCategoryRules SET SortOrder=" & SortArray(1, SortArrayPointer) & " WHERE (CategoryID=" & SortArray(0, SortArrayPointer) & ") AND (NewsletterIssueID=" & cp.Db.EncodeSQLNumber(IssueID) & ")"
                            Call cp.Db.ExecuteSQL(SQL)
                        Next
                    End If
                End If
                '
            End If
        End Sub
        '
        Friend Shared Function GetCategoryAccessString(cp As CPBaseClass, CategoryID As Integer) As String
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim SQL As String
            Dim Stream As String = ""
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
        Friend Shared Function GetArticleAccessString(cp As CPBaseClass, StoryID As Integer) As String
            '
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim SQL As String
            Dim Stream As String = ""
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
        Friend Shared Function HasAccess(cp As CPBaseClass, GroupString As String) As Boolean
            '
            Dim ListArray() As String
            Dim ListArrayCount As Integer
            Dim ListArrayPointer As Integer
            '
            If cp.User.IsContentManager("Newsletters") Then
                HasAccess = True
            Else
                If GroupString <> "" Then
                    If InStr(1, GroupString, ",", vbTextCompare) <> 0 Then
                        ListArray = Split(GroupString, ",", , vbTextCompare)
                        ListArrayCount = UBound(ListArray)
                        For ListArrayPointer = 0 To ListArrayCount
                            If cp.User.IsInGroup(cp.Content.GetRecordName("Groups", CInt(ListArray(ListArrayPointer)))) Then
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
        Private Shared Function PadValue(cp As CPBaseClass, Value As Integer, StringLenghth As Integer) As String
            Dim Counter As Integer
            Dim ValueLenghth As Integer
            Dim InnerValue As String
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
        Private Shared Function GetSortOrder(cp As CPBaseClass, CategoryID As Integer, IssueID As Integer) As String
            Dim cs As CPCSBaseClass = cp.CSNew()
            Dim Stream As String = ""
            '
            Call cs.Open("Newsletter Issue Category Rules", "(CategoryID=" & CategoryID & ") AND (NewsletterIssueID=" & IssueID & ")")
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
        Friend Shared Function verifyDefaultTemplateGetId(cp As CPBaseClass) As Integer
            Using cs As CPCSBaseClass = cp.CSNew()
                '
                ' -- try default template
                Call cs.Open("Newsletter Templates", "name='Default'")
                If Not cs.OK() Then
                    Call cs.Close()
                    Call cs.Insert("Newsletter Templates")
                    If cs.OK() Then
                        Call cs.SetField("name", "Default")
                    End If
                End If
                If cs.OK() Then
                    '
                    ' Use the default template in their Db already
                    If String.IsNullOrEmpty(Trim(cs.GetText("Template"))) Then
                        Call cs.SetField("Template", cp.WwwFiles.Read("newsletteraddon\Newsletter Template Default.html"))
                    End If
                    Return cs.GetInteger("ID")
                End If
                Call cs.Close()
                Return 0
            End Using
        End Function
        '
        Friend Shared Function verifyDefaultEmailTemplateGetId(cp As CPBaseClass) As Integer
            Using cs As CPCSBaseClass = cp.CSNew()
                '
                ' -- try default template
                Call cs.Open("Newsletter Templates", "name='Default Email'")
                If Not cs.OK() Then
                    Call cs.Close()
                    Call cs.Insert("Newsletter Templates")
                    If cs.OK() Then
                        Call cs.SetField("name", "Default Email")
                    End If
                End If
                If cs.OK() Then
                    '
                    ' Use the default template in their Db already
                    If String.IsNullOrEmpty(Trim(cs.GetText("Template"))) Then
                        Call cs.SetField("Template", cp.WwwFiles.Read("newsletteraddon\Newsletter Template Default Email.html"))
                    End If
                    Return cs.GetInteger("ID")
                End If
                Call cs.Close()
                Return 0
            End Using
        End Function
        '
        '===================================================================================================
        '   Wrap the content in a common wrapper if authoring is enabled
        '===================================================================================================
        '
        Public Shared Function GetEditWrapper(cp As CPBaseClass, Caption As String, Content As String) As String
            Return cp.Content.GetEditWrapper(Content)
        End Function
        '
        '===================================================================================================
        '   Wrap the content in a common wrapper if authoring is enabled
        '===================================================================================================
        '
        Public Shared Function GetAdminHintWrapper(cp As CPBaseClass, Content As String) As String
            Return cp.Html.adminHint(Content)
        End Function
        '
        Friend Shared Function encodeMinDate(source As Date) As Date
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
        Public Shared Function GetRandomInteger() As Integer
            Dim RandomLimit As Integer
            RandomLimit = 32767
            Randomize()
            GetRandomInteger = CInt(Rnd() * RandomLimit)
        End Function
        '
        '
        '
        Friend Shared Function isBlank(ByVal cp As CPBaseClass, ByVal source As String) As Boolean
            Dim returnBool As Boolean = False
            Try
                Dim test As String = source
                '
                If test.Length = 0 Then
                    returnBool = True
                ElseIf test.Length < 1000 Then
                    test = cp.Utils.ConvertHTML2Text(test)
                    test = test.Replace(vbLf, "")
                    test = test.Replace(vbCr, "")
                    test = test.Replace(vbTab, "")
                    test = test.Replace(" ", "")
                    returnBool = (test.Length = 0)
                End If
            Catch ex As Exception
                Call handleError(cp, ex, "isBlank")
            End Try
            Return returnBool
        End Function
    End Class
End Namespace
