﻿Namespace newsletter2
    Module newsletterModule
        Public Const DefaultEmailLinkSubject = "A link to NCA Currents newsletter"
        '
        Public Const RequestNameFormID = "formid"
        Public Const RequestNameIssueID = "issue"
        Public Const RequestNameIssuePageID = "issuepage"
        Public Const RequestNameChildPageID = "bid"
        '
        Public Const RequestNameEmailGroups = "Groups"
        '
        Public Const ContentNameNewsletters = "Newsletters"
        Public Const ContentNameNewsletterIssues = "Newsletter Issues"
        Public Const ContentNameNewsletterIssuePages = "Newsletter Issue Pages"
        Public Const ContentNameNewsLetterGroupRules = "Newsletter Group Rules"
        Public Const ContentNameMemberRules = "Member Rules"
        Public Const ContentNameGroupEmail = "Group Email"
        Public Const ContentNameGroups = "Groups"
        Public Const ContentNameEmailGroups = "Email Groups"
        Public Const ContentNameEmailTemplates = "Email Templates"
        Public Const ContentNameTopicRules = "Topic Rules"
        Public Const ContentNameTopics = "Topics"
        Public Const ContentNameIssueRules = "Newsletter Issue Category Rules"
        Public Const ContentNameIssueCategories = "Newsletter Issue Categories"
        '
        Public Const RequestNamePageNumber = "PageNumber"
        Public Const RequestNameRecordTop = "RecordTop"
        Public Const RequestNameNewsletterID = "nlid"
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
        Public Const PropertyAddOnVersion = "aoNewsletterVersion"
        Public Const VisitPropertyNewsletter = "SelectedNewsletter"
        '
        Public Const StringReplaceStart = "{{"
        Public Const StringReplaceEnd = "}}"
        '
        Public Const FormIssue = 100
        Public Const FormArchive = 200
        Public Const FormEmail = 300
        Public Const FormDetails = 400
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
        Public Const DefaultTemplate = "<TABLE style=""BORDER: black 1px solid; WIDTH: 100%; BORDER-COLLAPSE: collapse"" cellPadding=5><TBODY><TR><TD colspan=2 style=""BORDER: black 1px solid;"">Newsletter Banner<BR>(Edit this template to design your newsletter)</TD></TR><TR><TD style=""BORDER: black 1px solid; VERTICAL-ALIGN: top; WIDTH: 150px;""><AC type=""AGGREGATEFUNCTION"" name=""Newsletter-nav only"" querystring="""" acinstanceid=""{{ACID0}}""><br><img src=""/cclib/images/spacer.gif"" width=""150"" height=""1""></TD><TD width=""99%"" style=""BORDER: black 1px solid; VERTICAL-ALIGN: top; TEXT-ALIGN: left""><AC type=""AGGREGATEFUNCTION"" name=""Newsletter-body only"" querystring="""" acinstanceid=""{{ACID1}}""></TD></TR></TBODY></TABLE>"
        '
        Public Const NewsletterAddonGuid = "{B9EB288C-C2BD-4DDF-8C70-E44828C45BA0}"
        '
        Public Const RequestNameRefer = "EditReferer"
        Public ReferLink As String

    End Module
End Namespace