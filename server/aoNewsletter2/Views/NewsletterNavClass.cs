
using System;
using Contensive.Addons.Newsletter.Controllers;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Views {
    // 
    // 
    public class NewsletterNavClass {
        // 
        // =====================================================================================
        // common report for this class
        // =====================================================================================
        // 
        private void handleError(CPBaseClass cp, Exception ex, string @method) {
            try {
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterNavClass." + @method);
            } catch (Exception) {
                //
                // stop anything thrown from cp errorReport
                //
            }
        }
        //
        internal string GetNav(CPBaseClass cp, int issueid, int NewsletterID, bool isContentManager, int FormID, string newsNav, int currentIssueId) {
            string returnHtml = "";
            try {
                var layout = new BlockClass();
                var repeatItem = new BlockClass();
                // 
                var cs = cp.CSNew();
                var CS2 = cp.CSNew();
                var CSPointer = cp.CSNew();
                string ThisSQL;
                // Dim WorkingStoryId As Integer
                string NavSQL;
                string CategoryName;
                string PreviousCategoryName = "";
                var cn = new NewsletterController();
                string AccessString;
                int CategoryID;
                string QS;
                var ArticleCount = default(int);
                string newsNavStoryItem;
                string newsNavCategoryItem;
                // Dim storyCaption As String
                string repeatList = "";
                // 
                layout.load(newsNav);
                newsNavStoryItem = layout.getClassOuter("newsNavStoryItem");
                newsNavCategoryItem = layout.getClassOuter("newsNavCategoryItem");
                // 
                repeatItem.load(newsNavStoryItem);
                QS = cp.Doc.RefreshQueryString;
                QS = cp.Utils.ModifyQueryString(QS, Constants.RequestNameIssueID, issueid.ToString(), true);
                QS = cp.Utils.ModifyQueryString(QS, Constants.RequestNameFormID, Constants.FormCover, true);
                repeatItem.setClassInner("newsNavItemCaption", "Home");
                // repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                repeatList += repeatItem.getHtml().Replace("href=\"?\"", "href=\"?" + QS + "\"");
                // 
                NavSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder, NIC.Name AS CategoryName";
                NavSQL = NavSQL + " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR";
                NavSQL = NavSQL + " Where (NIC.ID = NIR.CategoryID)";
                NavSQL = NavSQL + " AND (NIR.NewsletterIssueID=" + issueid + ")";
                NavSQL = NavSQL + " AND (NIC.Active<>0)";
                NavSQL = NavSQL + " AND (NIR.Active<>0)";
                NavSQL = NavSQL + " ORDER BY NIR.SortOrder,NIC.name";
                // 
                cs.OpenSQL(NavSQL);
                if (cs.OK()) {
                    while (cs.OK()) {
                        CategoryID = cs.GetInteger("CategoryID");
                        CS2.Open(Constants.ContentNameNewsletterStories, "(CategoryID=" + CategoryID + ") AND (NewsletterID=" + issueid + ")", "SortOrder,id");
                        if (CS2.OK()) {
                            CategoryName = cs.GetText("CategoryName");
                            if ((CategoryName ?? "") != (PreviousCategoryName ?? "")) {
                                AccessString = NewsletterController.GetCategoryAccessString(cp, cs.GetInteger("CategoryID"));
                                if (!string.IsNullOrEmpty(AccessString)) {
                                    repeatList += "<AC type=\"AGGREGATEFUNCTION\" name=\"block text\" querystring=\"allowgroups=" + AccessString + "\">";
                                }
                                // 
                                repeatItem.load(newsNavCategoryItem);
                                repeatItem.setClassInner("newsNavItemCaption", CategoryName);
                                repeatList += repeatItem.getHtml();
                                // 
                                if (!string.IsNullOrEmpty(AccessString)) {
                                    repeatList += "<AC type=\"AGGREGATEFUNCTION\" name=\"block text end\" >";
                                }
                                PreviousCategoryName = CategoryName;
                            }
                            // 
                            while (CS2.OK()) {
                                // 
                                repeatList += getNavItem(cp, cn, CS2, newsNavStoryItem);
                                // Call repeatItem.Load(newsNavStoryItem)
                                // WorkingStoryId = CS2.GetInteger("ID")
                                // AccessString = NewsletterController.GetArticleAccessString(cp, WorkingStoryId)
                                // storyCaption = CS2.GetText("Name")
                                // Call repeatItem.SetClassInner("newsNavItemCaption", storyCaption)
                                // If AccessString <> "" Then
                                // repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>")
                                // End If
                                // QS = cp.Doc.RefreshQueryString
                                // QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, CStr(WorkingStoryId), True)
                                // QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                                // If AccessString <> "" Then
                                // repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                                // End If
                                // repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                                // 
                                ArticleCount = ArticleCount + 1;
                                CS2.GoNext();
                            }
                        }
                        CS2.Close();
                        // 
                        cs.GoNext();
                    }
                }
                cs.Close();
                // 
                cs.Open(Constants.ContentNameNewsletterStories, "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" + issueid + ")", "SortOrder,DateAdded");
                if (cs.OK()) {
                    if (ArticleCount > 0) {
                        // 
                        // This is a list of uncategorized articles following the categories -- give it a heading
                        // 
                        CategoryName = cp.Site.GetText("Newsletter Nav Caption Other Articles", "Other Articles");
                        repeatItem.load(newsNavCategoryItem);
                        repeatItem.setClassInner("newsNavItemCaption", CategoryName);
                        repeatList += repeatItem.getHtml();
                    }
                    while (cs.OK()) {
                        repeatList += getNavItem(cp, cn, cs, newsNavStoryItem);
                        // Call repeatItem.Load(newsNavStoryItem)
                        // WorkingStoryId = cs.GetInteger("ID")
                        // AccessString = NewsletterController.GetArticleAccessString(cp, WorkingStoryId)
                        // storyCaption = cs.GetText("Name")
                        // 'storyCaption = CS.GetEditLink() & CS.GetText("Name")
                        // If AccessString <> "" Then
                        // repeatItem.Prepend("<AC type=""AGGREGATEFUNCTION"" name=""block text"" querystring=""allowgroups=" & AccessString & """>")
                        // End If
                        // Call repeatItem.SetClassInner("newsNavItemCaption", storyCaption)
                        // If Not NewsletterController.isBlank(cp, cs.GetText("body")) Then
                        // 'If cs.GetBoolean("AllowReadMore") Then
                        // '
                        // ' link to the story page
                        // '
                        // QS = cp.Doc.RefreshQueryString
                        // QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, CStr(WorkingStoryId), True)
                        // QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormDetails, True)
                        // Else
                        // '
                        // ' link to the bookmark 'story#' on the cover
                        // '
                        // QS = "?" & cp.Doc.RefreshQueryString
                        // QS = cp.Utils.ModifyQueryString(QS, RequestNameStoryId, "", False)
                        // QS = cp.Utils.ModifyQueryString(QS, RequestNameFormID, FormCover, True)
                        // QS = QS & "#story" & WorkingStoryId
                        // End If
                        // If AccessString <> "" Then
                        // repeatItem.Append("<AC type=""AGGREGATEFUNCTION"" name=""block text end"" >")
                        // End If
                        // repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                        cs.GoNext();
                    }
                }
                cs.Close();
                // 
                // Link to Current Issues
                // 
                if (issueid != currentIssueId & currentIssueId != 0) {
                    QS = cp.Doc.RefreshQueryString;
                    QS = cp.Utils.ModifyQueryString(QS, Constants.RequestNameFormID, Constants.FormCover);
                    repeatItem.load(newsNavStoryItem);
                    repeatItem.setClassInner("newsNavItemCaption", cp.Site.GetText(Constants.SitePropertyCurrentIssue, "Current Issue"));
                    // repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                    repeatList += repeatItem.getHtml().Replace("href=\"?\"", "href=\"?" + QS + "\"");
                }
                // 
                // Display Archive Link if there are archive issues
                // can not just lookup issues that are not the issueid because if you are editing a future issue, the current issue shows up as an archive
                // 
                ThisSQL = "SELECT TOP 2 ID From NewsletterIssues WHERE active=1 and (PublishDate < { fn NOW() }) AND (NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ")";
                cs.OpenSQL(ThisSQL);
                if (cs.OK()) {
                    // 
                    // First one is the current issue
                    // 
                    cs.GoNext();
                    if (cs.OK()) {
                        // 
                        // If there are more then one published issues, the others are archive issues
                        // 
                        repeatItem.load(newsNavStoryItem);
                        repeatItem.setClassInner("newsNavItemCaption", cp.Site.GetText(Constants.SitePropertyIssueArchive, "Archives"));
                        QS = cp.Doc.RefreshQueryString;
                        // QS = cp.Utils.ModifyQueryString(QS, RequestNameNewsletterID, NewsletterID)
                        QS = cp.Utils.ModifyQueryString(QS, Constants.RequestNameFormID, Constants.FormArchive);
                        // repeatList &= repeatItem.GetHtml().Replace("?", "?" & QS)
                        repeatList += repeatItem.getHtml().Replace("href=\"?\"", "href=\"?" + QS + "\"");
                    }
                }
                cs.Close();
                // 
                layout.setClassInner("newsNavList", repeatList);
                // 
                returnHtml = layout.getHtml();
            } catch (Exception) {

            }
            return returnHtml;
        }
        // 
        private string GetArchiveLink(CPBaseClass cp, int newsletterId) {
            string GetArchiveLinkRet = default;
            string Stream = "";
            string qs = "";
            // 
            qs = cp.Doc.RefreshQueryString;
            qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormArchive);
            Stream += "<a class=\"caption\" href=\"?" + qs + "\">" + cp.Site.GetText(Constants.SitePropertyIssueArchive, "Archives") + "</a>";
            GetArchiveLinkRet = Stream;
            return GetArchiveLinkRet;
        }
        // 
        private string GetCurrentIssueLink(CPBaseClass cp) {
            string GetCurrentIssueLinkRet = default;
            string Stream = "";
            string qs = "";
            // 
            qs = cp.Doc.RefreshQueryString;
            qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormCover);
            Stream += "<a class=\"caption\" href=\"?" + qs + "\">" + cp.Site.GetText(Constants.SitePropertyCurrentIssue, "Current Issue") + "</a>";
            GetCurrentIssueLinkRet = Stream;
            return GetCurrentIssueLinkRet;
        }
        // 
        private string getNavItem(CPBaseClass cp, NewsletterController cn, CPCSBaseClass cs, string newsNavStoryItemLayout) {
            string returnHtml = "";
            try {
                var repeatItem = new BlockClass();
                int WorkingStoryId;
                string accessString;
                string storyCaption;
                string qs;
                // 
                repeatItem.load(newsNavStoryItemLayout);
                WorkingStoryId = cs.GetInteger("ID");
                accessString = NewsletterController.GetArticleAccessString(cp, WorkingStoryId);
                storyCaption = cs.GetText("Name");
                // storyCaption = CS.GetEditLink() & CS.GetText("Name")
                if (!string.IsNullOrEmpty(accessString)) {
                    repeatItem.load("<AC type=\"AGGREGATEFUNCTION\" name=\"block text\" querystring=\"allowgroups=" + accessString + "\">" + repeatItem.getHtml());
                }
                repeatItem.setClassInner("newsNavItemCaption", storyCaption);
                if (!NewsletterController.isBlank(cp, cs.GetText("body"))) {
                    // If cs.GetBoolean("AllowReadMore") Then
                    // 
                    // link to the story page
                    // 
                    qs = cp.Doc.RefreshQueryString;
                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameStoryId, WorkingStoryId.ToString(), true);
                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormStory, true);
                } else {
                    // 
                    // link to the bookmark 'story#' on the cover
                    // 
                    qs = cp.Doc.RefreshQueryString;
                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameStoryId, "", false);
                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormCover, true);
                    qs = qs + "#story" + WorkingStoryId;
                }
                if (!string.IsNullOrEmpty(accessString)) {
                    repeatItem.load(repeatItem.getHtml() + "<AC type=\"AGGREGATEFUNCTION\" name=\"block text end\" >");
                }
                // returnHtml = repeatItem.GetHtml().Replace("?", "?" & qs)
                returnHtml = repeatItem.getHtml().Replace("href=\"?\"", "href=\"?" + qs + "\"");
            } catch (Exception ex) {
                handleError(cp, ex, "getNavItem");
            }
            return returnHtml;
        }
        // 
    }
}