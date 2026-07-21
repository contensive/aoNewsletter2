using System;

namespace Contensive.Addons.Newsletter {

    static class Constants {
        // 
        /// <summary>
    /// newsletter version used in onInstall to migrate data. Site property saves the version last installed
    /// </summary>
        public const int newsletterVersion = 3;
        /// <summary>
    /// site property name for last installed version
    /// </summary>
        public const string spName_NewsletterVersion = "aoNewsletterVersion";
        public const string DefaultEmailLinkSubject = "A link to NCA Currents newsletter";
        // 
        public const string guidLayoutDefaultTemplate = "{24AB7755-C205-4C12-AE51-A58F885F1596}";
        public const string guidLayoutDefaultEmailTemplate = "{02FCBD4C-C610-493B-AA44-453DD3D8865A}";
        public const string guidLayoutDefaultIssueCover = "{5CBE8DAB-4CC0-4114-BADB-248D7929C486}";
        public const string guidLayoutDefaultStoryOverview = "{C10A769F-CB9B-4E69-AFCC-858852E90782}";
        public const string guidLayoutDefaultStoryBody = "{4A1A2A6B-E422-464E-B76F-5C254E31E1EF}";
        // guidLayoutDefaultStoryBody
        // 
        public const string RequestNameFormID = "formid";
        public const string RequestNameIssueID = "issue";
        public const string RequestNameStoryId = "storyId";
        public const string RequestNameChildPageID = "bid";
        // 
        public const string RequestNameEmailGroups = "Groups";
        // 
        public const string ContentNameNewsletters = "Newsletters";
        public const string ContentNameNewsletterIssues = "Newsletter Issues";
        public const string ContentNameNewsletterStories = "Newsletter Stories";
        // Public Const ContentNameNewsLetterGroupRules = "Newsletter Group Rules"
        public const string ContentNameMemberRules = "Member Rules";
        public const string ContentNameGroupEmail = "Group Email";
        public const string ContentNameGroups = "Groups";
        public const string ContentNameEmailGroups = "Email Groups";
        public const string ContentNameEmailTemplates = "Email Templates";
        public const string ContentNameTopicRules = "Topic Rules";
        public const string ContentNameTopics = "Topics";
        public const string ContentNameIssueRules = "Newsletter Issue Category Rules";
        public const string ContentNameIssueCategories = "Newsletter Categories";
        // 
        public const string RequestNamePageNumber = "PageNumber";
        public const string RequestNameRecordTop = "RecordTop";
        // Public Const RequestNameNewsletterID = "nlid"
        public const string RequestNameSortUp = "srtup";
        // 
        public const string FormButtonSend = " Send Email ";
        public const string FormButtonCreateEmail = " Create Email ";
        // 
        public const string TemplateName = "CHC Newsletter";
        public const string TemplateReplacementNav = "EmailnewsletterNavClass";
        public const string TemplateReplacementBody = "EmailNewsletterBody";
        public const string TemplateReplacementTitle = "EmailTitle";
        public const string TemplateReplacementPubDate = "EmailPublishDate";
        // 
        public const string EmailNewsFromAddress = "\"CHC News\"<chcnews@healthcharities.org>";
        // 
        public const string RequestNameMonthSelectd = "monthSelected";
        public const string RequestNameYearSelected = "yearSelected";
        public const string RequestNameButtonValue = "button";
        public const string FormButtonViewArchives = " Search ";
        public const string RequestNameSearchKeywords = "SearchKeywords";
        public const string FormButtonViewNewsLetter = " View ";
        // 
        public const string DefaultRecord = "Default";
        // 
        public const string VisitPropertyNewsletter = "SelectedNewsletter";
        // 
        public const string StringReplaceStart = "{{";
        public const string StringReplaceEnd = "}}";
        // 
        public const int FormCover = 100;
        public const int FormArchive = 200;
        public const int FormEmail = 300;
        public const int FormStory = 400;
        public const int FormSearch = 500;
        // 
        public const string PageNameArchives = "Newsletter Archives";
        public const string SitePropertyNoNewsletterArchives = "NoNewsletterArchives";
        // 
        public const string SitePropertyPageListCaption = "NewsletterPageListCaption";
        public const string SitePropertyIssueArchive = "NewsletterArchiveCaption";
        public const string SitePropertyCurrentIssue = "CurrentIssueCaption";
        public const string SitePropertyItemDelimiter = "NewsletterNaigationDelimiter";
        // '
        public const string ClassNameNavigationTable = "NewsletterNavigationTable";
        public const string ClassNameGroupHeader = "Caption";
        public const string ClassNamePageList = "PageList";
        // 
        public const string ModeHorizontal = "Horizontal";
        public const string ModeVertical = "Vertical";
        // 
        // Public Const DefaultTemplate = "<TABLE style=""BORDER: black 1px solid; WIDTH: 100%; BORDER-COLLAPSE: collapse"" cellPadding=5><TBODY><TR><TD colspan=2 style=""BORDER: black 1px solid;"">Newsletter Banner<BR>(Edit this template to design your newsletter)</TD></TR><TR><TD style=""BORDER: black 1px solid; VERTICAL-ALIGN: top; WIDTH: 150px;""><AC type=""AGGREGATEFUNCTION"" name=""Newsletter-nav only"" querystring="""" acinstanceid=""{{ACID0}}""><br><img src=""/cclib/images/spacer.gif"" width=""150"" height=""1""></TD><TD width=""99%"" style=""BORDER: black 1px solid; VERTICAL-ALIGN: top; TEXT-ALIGN: left""><AC type=""AGGREGATEFUNCTION"" name=""Newsletter-body only"" querystring="""" acinstanceid=""{{ACID1}}""></TD></TR></TBODY></TABLE>"
        // 
        public const string NewsletterAddonGuid = "{B9EB288C-C2BD-4DDF-8C70-E44828C45BA0}";
        //
        // -- Portal integration constants
        //
        public const string guidPortalShare = "{e4d011e9-9f3b-4f7e-8ec3-f4fcc2a20455}";
        public const string guidPortalFeatureNewsletterList = "{458709B9-FEBA-475E-9F17-41B0D912A798}";
        public const string guidPortalFeatureNewsletterIssueList = "{369607B7-7FAA-424F-A06F-E2E8652EE11B}";
        public const string guidAddonNewsletterList = "{951FA2BE-B559-495C-9CB6-1D65E1A074A0}";
        public const string guidAddonNewsletterIssueList = "{C24A2D81-395B-4A47-BD7D-0904330E28A1}";
        //
        // -- Portal request/form constants
        //
        public const string rnNewsletterId = "newsletterId";
        public const string rnButton = "button";
        public const string rnSrcFormId = "srcFormId";
        public const string rnDstFeatureGuid = "dstFeatureGuid";
        public const int formIdNewsletterList = 1;
        public const int formIdNewsletterIssueList = 2;
        public const string buttonAdd = "Add";
        public const string buttonBack = "Back";
        //
        public const string RequestNameRefer = "EditReferer";
        public static string ReferLink = "";
        // 
        public static string openRecord(BaseClasses.CPBaseClass cp, ref BaseClasses.CPCSBaseClass cs, string contentName, int recordID, string fieldList = "") {
            string s = "";
            // 
            try {
                cs.Open(contentName, "id=" + recordID, "", true, fieldList);
            } catch (Exception ex) {
                try {
                    cp.Site.ErrorReport(ex, "error in newsletter2.newsletterCommonClass.openRecord");
                } catch (Exception) {
                }
            }
            // 
            return s;
        }
        // 
    }
}