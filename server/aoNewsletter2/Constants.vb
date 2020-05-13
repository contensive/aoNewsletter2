
Module Constants
    '
    ''' <summary>
    ''' newsletter version used in onInstall to migrate data. Site property saves the version last installed
    ''' </summary>
    Public Const newsletterVersion As Integer = 3
    ''' <summary>
    ''' site property name for last installed version
    ''' </summary>
    Public Const spName_NewsletterVersion = "aoNewsletterVersion"
    Public Const DefaultEmailLinkSubject = "A link to NCA Currents newsletter"
    '
    Public Const guidLayoutDefaultTemplate = "{24AB7755-C205-4C12-AE51-A58F885F1596}"
    Public Const guidLayoutDefaultEmailTemplate = "{02FCBD4C-C610-493B-AA44-453DD3D8865A}"
    Public Const guidLayoutDefaultIssueCover = "{5CBE8DAB-4CC0-4114-BADB-248D7929C486}"
    Public Const guidLayoutDefaultStoryOverview = "{C10A769F-CB9B-4E69-AFCC-858852E90782}"
    Public Const guidLayoutDefaultStoryBody = "{4A1A2A6B-E422-464E-B76F-5C254E31E1EF}"
    'guidLayoutDefaultStoryBody
    '
    Public Const RequestNameFormID = "formid"
    Public Const RequestNameIssueID = "issue"
    Public Const RequestNameStoryId = "storyId"
    Public Const RequestNameChildPageID = "bid"
    '
    Public Const RequestNameEmailGroups = "Groups"
    '
    Public Const ContentNameNewsletters = "Newsletters"
    Public Const ContentNameNewsletterIssues = "Newsletter Issues"
    Public Const ContentNameNewsletterStories = "Newsletter Stories"
    'Public Const ContentNameNewsLetterGroupRules = "Newsletter Group Rules"
    Public Const ContentNameMemberRules = "Member Rules"
    Public Const ContentNameGroupEmail = "Group Email"
    Public Const ContentNameGroups = "Groups"
    Public Const ContentNameEmailGroups = "Email Groups"
    Public Const ContentNameEmailTemplates = "Email Templates"
    Public Const ContentNameTopicRules = "Topic Rules"
    Public Const ContentNameTopics = "Topics"
    Public Const ContentNameIssueRules = "Newsletter Issue Category Rules"
    Public Const ContentNameIssueCategories = "Newsletter Categories"
    '
    Public Const RequestNamePageNumber = "PageNumber"
    Public Const RequestNameRecordTop = "RecordTop"
    'Public Const RequestNameNewsletterID = "nlid"
    Public Const RequestNameSortUp = "srtup"
    '
    Public Const FormButtonSend = " Send Email "
    Public Const FormButtonCreateEmail = " Create Email "
    '
    Public Const TemplateName = "CHC Newsletter"
    Public Const TemplateReplacementNav = "EmailnewsletterNavClass"
    Public Const TemplateReplacementBody = "EmailNewsletterBody"
    Public Const TemplateReplacementTitle = "EmailTitle"
    Public Const TemplateReplacementPubDate = "EmailPublishDate"
    '
    Public Const EmailNewsFromAddress = """CHC News""<chcnews@healthcharities.org>"
    '
    Public Const RequestNameMonthSelectd = "monthSelected"
    Public Const RequestNameYearSelected = "yearSelected"
    Public Const RequestNameButtonValue = "button"
    Public Const FormButtonViewArchives = " Search "
    Public Const RequestNameSearchKeywords = "SearchKeywords"
    Public Const FormButtonViewNewsLetter = " View "
    '
    Public Const DefaultRecord = "Default"
    '
    Public Const VisitPropertyNewsletter = "SelectedNewsletter"
    '
    Public Const StringReplaceStart = "{{"
    Public Const StringReplaceEnd = "}}"
    '
    Public Const FormCover = 100
    Public Const FormArchive = 200
    Public Const FormEmail = 300
    Public Const FormStory = 400
    Public Const FormSearch = 500
    '
    Public Const PageNameArchives = "Newsletter Archives"
    Public Const SitePropertyNoNewsletterArchives = "NoNewsletterArchives"
    '
    Public Const SitePropertyPageListCaption = "NewsletterPageListCaption"
    Public Const SitePropertyIssueArchive = "NewsletterArchiveCaption"
    Public Const SitePropertyCurrentIssue = "CurrentIssueCaption"
    Public Const SitePropertyItemDelimiter = "NewsletterNaigationDelimiter"
    ''
    Public Const ClassNameNavigationTable = "NewsletterNavigationTable"
    Public Const ClassNameGroupHeader = "Caption"
    Public Const ClassNamePageList = "PageList"
    '
    Public Const ModeHorizontal = "Horizontal"
    Public Const ModeVertical = "Vertical"
    '
    'Public Const DefaultTemplate = "<TABLE style=""BORDER: black 1px solid; WIDTH: 100%; BORDER-COLLAPSE: collapse"" cellPadding=5><TBODY><TR><TD colspan=2 style=""BORDER: black 1px solid;"">Newsletter Banner<BR>(Edit this template to design your newsletter)</TD></TR><TR><TD style=""BORDER: black 1px solid; VERTICAL-ALIGN: top; WIDTH: 150px;""><AC type=""AGGREGATEFUNCTION"" name=""Newsletter-nav only"" querystring="""" acinstanceid=""{{ACID0}}""><br><img src=""/cclib/images/spacer.gif"" width=""150"" height=""1""></TD><TD width=""99%"" style=""BORDER: black 1px solid; VERTICAL-ALIGN: top; TEXT-ALIGN: left""><AC type=""AGGREGATEFUNCTION"" name=""Newsletter-body only"" querystring="""" acinstanceid=""{{ACID1}}""></TD></TR></TBODY></TABLE>"
    '
    Public Const NewsletterAddonGuid = "{B9EB288C-C2BD-4DDF-8C70-E44828C45BA0}"
    '
    Public Const RequestNameRefer = "EditReferer"
    Public ReferLink As String
    '
    Public Function openRecord(ByVal cp As Global.Contensive.BaseClasses.CPBaseClass, ByRef cs As Global.Contensive.BaseClasses.CPCSBaseClass, ByVal contentName As String, ByVal recordID As Integer, Optional ByVal fieldList As String = "") As String
        Dim s As String = ""
        '
        Try
            cs.Open(contentName, "id=" & recordID, "", True, fieldList)
        Catch ex As Exception
            Try
                cp.Site.ErrorReport(ex, "error in newsletter2.newsletterCommonClass.openRecord")
            Catch errObj As Exception
            End Try
        End Try
        '
        Return s
    End Function
    '
End Module
