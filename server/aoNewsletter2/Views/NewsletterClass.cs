
using System;
using System.Collections.Generic;
using Contensive.Addons.Newsletter.Controllers;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Views {
    public class NewsletterClass {
        // 
        // =====================================================================================
        /// <summary>
        /// Newsletter Addon Interface
        /// </summary>
        /// <param name="CP"></param>
        /// <returns></returns>
        public object getLegacyNewsletter(CPBaseClass CP, int NewsletterID, int currentIssueID) {
            string returnHtml = "";
            try {
                string refreshQueryString = "";
                // 
                var layout = new BlockClass();
                string newsBody = "";
                string newsNav = "";
                // 
                string EditLink;
                string Controls;
                string UnpublishedIssueList;
                bool BuildDefault;
                var IssueID = default(int);
                int storyID;
                var cn = new NewsletterController();
                var cs = CP.CSNew();
                NewsletterBodyClass Body;
                NewsletterNavClass nav;
                var TemplateID = default(int);
                int FormID;
                int EmailID;
                string TemplateCopy = "";
                string qs;
                string ButtonValue;

                bool isManager;
                string ReferLink;
                string currentLink = "";
                bool isContentManager = CP.User.IsContentManager("newsletters");
                string itemLayout = "";
                string itemLayoutStory = "";
                string itemLayoutCategory = "";
                bool isEditing = CP.User.IsEditing();
                string ItemList = "";

                string footerAdBanners = "";
                string itemLayoutAdBanners = "";
                string sponsor = "";
                var publishDate = DateTime.MinValue;
                string tagLine = "";
                string mastheadFilename = "";
                string footerFilename = "";
                var problemList = new List<string>();
                // 
                refreshQueryString = CP.Doc.RefreshQueryString;
                // 
                currentLink = CP.Request.Protocol + CP.Site.DomainPrimary + CP.Request.PathPage + "?" + refreshQueryString;
                ReferLink = Constants.RequestNameRefer + "=" + CP.Utils.EncodeRequestVariable(CP.Utils.ModifyLinkQueryString(currentLink, Constants.RequestNameRefer, ""));
                isManager = CP.User.IsContentManager("Newsletters");
                // 
                BuildDefault = CP.Doc.GetBoolean("BuildDefault");
                FormID = CP.Doc.GetInteger(Constants.RequestNameFormID);
                storyID = CP.Doc.GetInteger(Constants.RequestNameStoryId);
                if (storyID == 0) {
                    // 
                    // No page given, use the QS for the Issue, or get current
                    // 
                    CP.Site.TestPoint("GetIssueID call 4, NewsletterID=" + NewsletterID);
                    IssueID = NewsletterController.GetIssueID(CP, NewsletterID, currentIssueID);
                } else {
                    // 
                    // PageID given, get Issue from PageID (and check against Newsletter)
                    // 
                    cs.Open(Constants.ContentNameNewsletterStories, "(id=" + storyID + ")");
                    if (cs.OK()) {
                        IssueID = cs.GetInteger("NewsletterID");
                    }
                    cs.Close();
                    // 
                    cs.Open(Constants.ContentNameNewsletterIssues, "active=1 and (id=" + IssueID + ")and(Newsletterid=" + NewsletterID + ")");
                    if (!cs.OK()) {
                        // 
                        // Bad Issue, reset to current issue of current newsletter
                        // 
                        CP.Site.TestPoint("GetIssueID call 5, NewsletterID=" + NewsletterID);
                        IssueID = NewsletterController.GetIssueID(CP, NewsletterID, currentIssueID);
                        storyID = 0;
                        FormID = Constants.FormCover;
                    }
                    cs.Close();
                }
                CP.Visit.SetProperty(Constants.VisitPropertyNewsletter, NewsletterID + "." + IssueID + "." + storyID + "." + FormID);
                // 
                CP.Site.TestPoint("PageClass NLID: " + NewsletterID);
                // 
                NewsletterController.SortCategoriesByIssue(CP, IssueID);
                // 
                if (isManager & FormID == Constants.FormEmail) {
                    // 
                    // create email version -- use Print Version to block edit links
                    // 
                    EmailID = CreateEmailGetID(CP, IssueID, NewsletterID, refreshQueryString, currentIssueID);
                    CP.Response.Redirect(CP.Site.GetText("adminUrl") + "?cid=" + CP.Content.GetID(Constants.ContentNameGroupEmail) + "&id=" + EmailID + "&af=4");
                    returnHtml = "";
                } else if (FormID == Constants.FormEmail) {
                    // 
                    // Not administrators
                    // 
                    CP.UserError.Add("Only administrators can use the Create Email feature.");
                    FormID = Constants.FormCover;
                } else {
                    // 
                    // Create the Newsletter
                    // 
                    if (IssueID == 0) {
                        // 
                        // There are no current issues, diplay a message and tell the admin what to do next
                        // 
                        returnHtml = "<p>There are currently no published issues of this newsletter</p>";
                    } else {
                        if (NewsletterID != 0) {
                            Constants.openRecord(CP, ref cs, "Newsletters", NewsletterID, "StylesFilename,TemplateID,mastheadFilename,footerFilename");
                            if (cs.OK()) {
                                TemplateID = cs.GetInteger("TemplateID");
                                mastheadFilename = cs.GetText("mastheadFilename");
                                footerFilename = cs.GetText("footerFilename");
                            }
                            cs.Close();
                            // 
                            if (TemplateID != 0) {
                                Constants.openRecord(CP, ref cs, "Newsletter Templates", TemplateID, "Template");
                                if (!cs.OK()) {
                                    // 
                                    // template set, but the ID is bad
                                    // 
                                    TemplateID = 0;
                                } else {
                                    TemplateCopy = cs.GetText("Template");
                                    if (string.IsNullOrEmpty(TemplateCopy)) {
                                        // 
                                        // template set, but the copy is empty
                                        // 
                                        TemplateID = 0;
                                    }
                                }
                                cs.Close();
                            }
                            // 
                            if (TemplateID == 0) {
                                TemplateID = NewsletterController.verifyDefaultTemplateGetId(CP);
                                if (TemplateID != 0) {
                                    Constants.openRecord(CP, ref cs, "Newsletters", IssueID);
                                    if (cs.OK()) {
                                        cs.SetField("TemplateID", TemplateID.ToString());
                                    }
                                    cs.Close();
                                }
                            }
                            // 
                            if (TemplateID > 0) {
                                Constants.openRecord(CP, ref cs, "Newsletter Templates", TemplateID);
                                if (cs.OK()) {
                                    EditLink = cs.GetEditLink();
                                    TemplateCopy = cs.GetText("Template");
                                }
                                cs.Close();
                            }
                        }

                        // 
                        // Process forms
                        // 
                        ButtonValue = CP.Doc.GetText("Button");
                        switch (FormID) {
                            case Constants.FormArchive: {
                                    switch (ButtonValue ?? "") {
                                        case Constants.FormButtonViewNewsLetter: {
                                                // 
                                                // Archive form pressing the view button
                                                // 
                                                FormID = Constants.FormCover;
                                                break;
                                            }
                                    }

                                    break;
                                }
                        }
                        // 
                        // Dispay the form
                        // 
                        layout.load(TemplateCopy);
                        // 
                        // -- masthead image
                        if (!string.IsNullOrEmpty(mastheadFilename)) {
                            mastheadFilename = Uri.EscapeUriString(mastheadFilename);
                            layout.setClassInner("newsHeaderMasthead", "<img src=\"" + CP.Http.CdnFilePathPrefix + mastheadFilename + "\" style=\"width:100%\" class=\"banner\" />");
                        }
                        // 
                        // -- footer image
                        if (!string.IsNullOrEmpty(footerFilename)) {
                            footerFilename = Uri.EscapeUriString(footerFilename);
                            layout.setClassInner("newsFooterMasthead", "<img src=\"" + CP.Http.CdnFilePathPrefix + footerFilename + "\" style=\"width:100%\" class=\"banner\" />");
                        }
                        // 
                        nav = new NewsletterNavClass();
                        newsNav = layout.getClassInner("newsNav");
                        // 
                        Body = new NewsletterBodyClass();
                        switch (FormID) {
                            case Constants.FormSearch: {
                                    itemLayout = layout.getClassOuter("newsSearchListItem");
                                    ItemList = Body.GetSearchItemList(CP, cn, ButtonValue, IssueID, refreshQueryString, itemLayout);
                                    itemLayoutAdBanners = layout.getClassOuter("adBannerItem");
                                    layout.setClassOuter("newsSearchList", ItemList);
                                    layout.setClassInner("newsArchive", "");
                                    layout.setClassOuter("newsBody", "");
                                    layout.setClassOuter("newsCover", "");
                                    layout.setClassOuter("emailLinkToWeb", "");
                                    layout.setClassOuter("newsIssueCaption", "");
                                    layout.setClassInner("newsIssueSponsor", sponsor);
                                    layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString());
                                    if (string.IsNullOrEmpty(tagLine)) {
                                        layout.setClassOuter("newsletterTagLine", "");
                                    } else {
                                        layout.setClassInner("newsletterTagLine", tagLine);
                                    }
                                    newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID);
                                    break;
                                }
                            case Constants.FormArchive: {
                                    // 
                                    // 
                                    // 
                                    string searchForm = "";
                                    searchForm += "<div>";
                                    searchForm += CP.Html.InputText(Constants.RequestNameSearchKeywords);
                                    searchForm += " <input type=\"submit\" id=\"js-ArchiveIssuesSubmit\" name=\"Button\" value=\" Search \"> "; // CP.Html.Button(FormButtonViewArchives, FormButtonViewArchives)
                                    searchForm += "</div>";
                                    searchForm = CP.Html.Form(searchForm, "", "", "", CP.Utils.ModifyQueryString(refreshQueryString, Constants.RequestNameFormID, Constants.FormArchive.ToString()));

                                    layout.setClassInner("newsArchiveSearch", searchForm);
                                    // 
                                    // 
                                    // 
                                    itemLayout = layout.getClassOuter("newsArchiveListItem");
                                    ItemList = Body.GetArchiveItemList(CP, cn, ButtonValue, currentIssueID, refreshQueryString, itemLayout, NewsletterID);
                                    itemLayoutAdBanners = layout.getClassOuter("adBannerItem");
                                    layout.setClassInner("newsArchiveList", ItemList);
                                    layout.setClassOuter("newsBody", "");
                                    layout.setClassOuter("newsCover", "");
                                    layout.setClassOuter("newsSearch", "");
                                    layout.setClassOuter("emailLinkToWeb", "");
                                    layout.setClassOuter("newsIssueCaption", "");
                                    layout.setClassInner("newsIssueSponsor", sponsor);
                                    layout.setClassInner("newsIssuePublishDate", "");
                                    if (string.IsNullOrEmpty(tagLine)) {
                                        layout.setClassOuter("newsletterTagLineRow", "");
                                    } else {
                                        layout.setClassInner("newsletterTagLine", "");
                                    }
                                    newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID);
                                    break;
                                }
                            case Constants.FormStory: {
                                    newsBody = layout.getClassInner("newsBody");
                                    if (string.IsNullOrEmpty(newsBody.Trim())) {
                                        problemList.Add("The newsletter template does not contain a class with 'newsBody', required for a detail page.");
                                    }
                                    newsBody = Body.GetStory(CP, cn, storyID, IssueID, refreshQueryString, newsBody, isEditing);
                                    Constants.openRecord(CP, ref cs, "Newsletter Issues", IssueID);
                                    if (cs.OK()) {
                                        sponsor = cs.GetText("sponsor");
                                        tagLine = cs.GetText("tagLine");
                                        publishDate = cs.GetDate("publishDate");
                                    }
                                    cs.Close();
                                    itemLayoutAdBanners = layout.getClassOuter("adBannerItem");
                                    layout.setClassInner("newsBody", newsBody);
                                    layout.setClassOuter("newsArchive", "");
                                    layout.setClassOuter("newsCover", "");
                                    layout.setClassOuter("newsSearch", "");
                                    layout.setClassOuter("emailLinkToWeb", "");
                                    layout.setClassInner("newsIssueCaption", CP.Content.GetRecordName(Constants.ContentNameNewsletterIssues, IssueID));
                                    layout.setClassInner("newsIssueSponsor", sponsor);
                                    layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString());
                                    if (string.IsNullOrEmpty(tagLine)) {
                                        layout.setClassOuter("newsletterTagLineRow", "");
                                    } else {
                                        layout.setClassInner("newsletterTagLine", tagLine);
                                    }
                                    newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID);
                                    break;
                                }

                            default: {
                                    // 
                                    // -- Form Cover
                                    FormID = Constants.FormCover;
                                    itemLayoutStory = layout.getClassOuter("newsCoverStoryItem");
                                    itemLayoutAdBanners = layout.getClassOuter("adBannerItem");
                                    itemLayoutCategory = layout.getClassOuter("newsCoverCategoryItem");
                                    ItemList = Body.GetCoverContent(CP, IssueID, storyID, refreshQueryString, FormID, itemLayoutStory, itemLayoutCategory, isEditing, ref sponsor, ref publishDate, ref tagLine);
                                    // 
                                    // add footer ad banner(s)
                                    // 
                                    if (cs.Open("newsletter Issues", "id=" + IssueID)) {
                                        string adBanner;
                                        string adBannerLink;
                                        int bannerLayoutId;
                                        int adBannerRowCnt = 1;
                                        int adBannerColumnCnt = 1;
                                        int pxColumnSpace = 0;
                                        int pxRowSpace = 0;
                                        // 
                                        bannerLayoutId = cs.GetInteger("bannerLayoutId");
                                        if (bannerLayoutId > 0) {
                                            var csLayout = CP.CSNew();
                                            if (csLayout.Open("Newsletter Ad Banner Layouts", "id=" + bannerLayoutId)) {
                                                adBannerRowCnt = csLayout.GetInteger("rowCnt");
                                                adBannerColumnCnt = csLayout.GetInteger("columnCnt");
                                                pxColumnSpace = csLayout.GetInteger("pxColumnSpace");
                                                pxRowSpace = csLayout.GetInteger("pxRowSpace");
                                            }
                                            csLayout.Close();
                                        }

                                        for (int rowPtr = 0, loopTo = adBannerRowCnt - 1; rowPtr <= loopTo; rowPtr++) {
                                            if (pxRowSpace > 0 & rowPtr > 0) {
                                                footerAdBanners += @"<img src=""\cclib\images\spacer.gif"" width=""10"" height=""" + pxRowSpace.ToString() + "\" style=\"height:" + pxRowSpace.ToString() + "px\">";
                                            }
                                            footerAdBanners += "<div class=\"newsletterAdvertisementRow\">";
                                            for (int columnPtr = 0, loopTo1 = adBannerColumnCnt - 1; columnPtr <= loopTo1; columnPtr++) {
                                                if (pxColumnSpace > 0 & columnPtr > 0) {
                                                    footerAdBanners += @"<img src=""\cclib\images\spacer.gif"" width=""" + pxColumnSpace.ToString() + "\" height=\"10\" style=\"width:" + pxColumnSpace.ToString() + "px\">";
                                                }
                                                int adPtr = rowPtr * adBannerColumnCnt + columnPtr;
                                                adBanner = cs.GetText("adBanner" + adPtr);
                                                if (!string.IsNullOrEmpty(adBanner)) {
                                                    adBannerLink = cs.GetText("adBannerLink" + adPtr);
                                                    if (string.IsNullOrEmpty(adBannerLink)) {
                                                        adBanner = Uri.EscapeUriString(adBanner);
                                                        footerAdBanners += "<img src=\"" + CP.Http.CdnFilePathPrefix + adBanner + "\">";
                                                    } else {
                                                        if (adBannerLink.IndexOf("://") < 0) {
                                                            adBannerLink = "http://" + adBannerLink;
                                                        }
                                                        adBanner = Uri.EscapeUriString(adBanner);
                                                        adBannerLink = Uri.EscapeUriString(adBannerLink);
                                                        footerAdBanners += "<a href=\"" + adBannerLink + "\" target=\"_blank\"><img src=\"" + CP.Http.CdnFilePathPrefix + adBanner + "\"></a>";
                                                    }
                                                }
                                            }
                                            footerAdBanners += "</div>";
                                        }
                                        // '
                                        // adBanner = cs.GetText("adBanner2")
                                        // If (Not String.IsNullOrEmpty(adBanner)) Then
                                        // footerAdBanners &= adBanner
                                        // End If
                                        // '
                                        // adBanner = cs.GetText("adBanner3")
                                        // If (Not String.IsNullOrEmpty(adBanner)) Then
                                        // footerAdBanners &= adBanner
                                        // End If
                                        // '
                                        // adBanner = cs.GetText("adBanner4")
                                        // If (Not String.IsNullOrEmpty(adBanner)) Then
                                        // footerAdBanners &= adBanner
                                        // End If
                                        // '
                                        // adBanner = cs.GetText("adBanner5")
                                        // If (Not String.IsNullOrEmpty(adBanner)) Then
                                        // footerAdBanners &= adBanner
                                        // End If
                                        // '
                                        // adBanner = cs.GetText("adBanner6")
                                        // If (Not String.IsNullOrEmpty(adBanner)) Then
                                        // footerAdBanners &= adBanner
                                        // End If
                                        // 
                                    }
                                    cs.Close();
                                    if (!string.IsNullOrEmpty(footerAdBanners)) {
                                        var adBannerLayout2 = new BlockClass();
                                        adBannerLayout2.load(footerAdBanners);
                                        adBannerLayout2.setClassInner("newsletterAdvertisements", footerAdBanners);
                                        ItemList += adBannerLayout2.getHtml();
                                    }
                                    layout.setClassInner("newsCoverList", ItemList);
                                    layout.setClassOuter("newsArchive", "");
                                    layout.setClassOuter("newsBody", "");
                                    layout.setClassOuter("newsSearch", "");
                                    layout.setClassOuter("emailLinkToWeb", "");
                                    layout.setClassInner("newsIssueCaption", CP.Content.GetRecordName(Constants.ContentNameNewsletterIssues, IssueID));
                                    layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString());
                                    if (string.IsNullOrWhiteSpace(sponsor)) {
                                        layout.setClassOuter("newsIssueSponsor", "");
                                    } else {
                                        layout.setClassInner("newsIssueSponsor", sponsor);
                                    }
                                    if (string.IsNullOrEmpty(tagLine)) {
                                        layout.setClassOuter("newsletterTagLineRow", "");
                                    } else {
                                        layout.setClassInner("newsletterTagLine", tagLine);
                                    }
                                    newsNav = nav.GetNav(CP, IssueID, NewsletterID, isContentManager, FormID, newsNav, currentIssueID);
                                    break;
                                }
                        }
                        layout.setClassInner("newsNav", newsNav);
                        // 
                        // Add archive link
                        // 
                        string newsArchiveLink = layout.getClassInner("newsArchiveLink");
                        newsArchiveLink = newsArchiveLink.Replace("#", CP.Utils.ModifyLinkQueryString(currentLink, "formId", Constants.FormArchive.ToString()));
                        layout.setClassInner("newsArchiveLink", newsArchiveLink);
                        // 
                        returnHtml = layout.getHtml();
                    }
                    // 
                    // List Unpublished issues for admins
                    // 
                    if (isEditing) {
                        // 
                        // -- wrap in issue edit
                        returnHtml = CP.Content.GetEditLink("newsletter issues", currentIssueID) + CP.Content.GetEditWrapper(returnHtml);
                        // 
                        // Controls
                        // 
                        Controls = "";
                        qs = refreshQueryString;
                        if (!string.IsNullOrEmpty(qs)) {
                            qs = qs + "&";
                        } else {
                            qs = qs + "?";
                        }
                        if (problemList.Count > 0) {
                            string controlItems = "";
                            foreach (string problem in problemList) {
                                controlItems += CP.Html.li(problem);
                            }
                            Controls = Controls + "<h3>Problems Found on this Page</h3>";
                            Controls += CP.Html.ul(controlItems);
                        }
                        if (IssueID != 0) {
                            // 
                            // For this issue
                            // 
                            Controls = Controls + "<h3>For this Issue</h3><ul>";
                            Controls = Controls + "<li><div class=\"AdminLink\"><a href = \"" + CP.Site.GetText("adminUrl") + "?cid=" + CP.Content.GetID(Constants.ContentNameNewsletterStories) + "&af=4&aa=2&ad=1&wc=" + CP.Utils.EncodeRequestVariable("NewsletterID=" + IssueID) + "&" + ReferLink + "\">Add a new story</a></div></li>";
                            Controls = Controls + "<li><div class=\"AdminLink\"><a href = \"" + CP.Site.GetText("adminUrl") + "?cid=" + CP.Content.GetID(Constants.ContentNameNewsletterIssues) + "&af=4&id=" + IssueID + "&" + ReferLink + "\">Edit this issue</a></div></li>";
                            if (CP.Request.PathPage.IndexOf("/admin", StringComparison.OrdinalIgnoreCase) >= 0 | ((CP.Site.GetText("adminUrl") ?? "").ToLowerInvariant() ?? "") == ((CP.Request.PathPage ?? "").ToLowerInvariant() ?? "")) {
                                Controls = Controls + "<li><div class=\"AdminLink\">Create&nbsp;email&nbsp;version (not available from admin site)</div></li>";
                            } else {
                                qs = CP.Doc.RefreshQueryString;
                                qs = CP.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormEmail.ToString());
                                qs = CP.Utils.ModifyQueryString(qs, Constants.RequestNameIssueID, IssueID.ToString());
                                Controls = Controls + "<li><div class=\"AdminLink\"><a href=\"?" + qs + "\">Create&nbsp;email&nbsp;version</a></div></li>";
                            }
                            Controls = Controls + "</ul>";
                        }
                        if (NewsletterID != 0) {
                            // 
                            // For this newsletter
                            // 
                            Controls = Controls + "<h3>For this Newsletter</h3><ul>";
                            Controls = Controls + "<li><div class=\"AdminLink\"><a href = \"" + CP.Site.GetText("adminUrl") + "?cid=" + CP.Content.GetID(Constants.ContentNameNewsletterIssues) + "&wl0=newsletterid&wr0=" + NewsletterID + "&af=4&aa=2&ad=1&" + "&" + ReferLink + "\">Add a new issue</a></div></li>";
                            Controls = Controls + "<li><div class=\"AdminLink\"><a href = \"" + CP.Site.GetText("adminUrl") + "?cid=" + CP.Content.GetID(Constants.ContentNameNewsletters) + "&id=" + NewsletterID + "&af=4&aa=2&ad=1&" + "&" + ReferLink + "\">Edit this newsletter</a></div></li>";
                            Controls = Controls + "</ul>";
                            // 
                            // Search for unpublished versions
                            // 
                            UnpublishedIssueList = NewsletterController.GetUnpublishedIssueList(CP, NewsletterID, cn);
                            if (!string.IsNullOrEmpty(UnpublishedIssueList)) {
                                Controls = Controls + "<h3>Unpublished issues for this Newsletter</h3>";
                                Controls = Controls + UnpublishedIssueList;
                            }
                        }
                        // 
                        // General Controls
                        // 
                        Controls = Controls + "<h3>General</h3><ul>";
                        Controls = Controls + "<li><div class=\"AdminLink\"><a href = \"" + CP.Site.GetText("adminUrl") + "?cid=" + CP.Content.GetID(Constants.ContentNameIssueCategories) + "&" + ReferLink + "\">Edit categories</a></div></li>";
                        // Controls = Controls & "<li><div class=""AdminLink""><a href = """ & CP.Site.GetText("adminUrl") & "?cid=" & CP.Content.GetID(ContentNameNewsletters) & "&af=4&" & "&" & ReferLink & """>Add a new newsletter</a></div></li>"
                        Controls = Controls + "</ul>";
                        // 
                        // instructions
                        // 
                        Controls = Controls + "<P>This addon can control one or many different newsletters on your site. For instance you may have a newsletter about site news and another about industry news. Each newsletter can have many issues. For instance, Site News may have a new issue every quarter, Industry News may have a new issue every month. Each issue can have many stories. The newsletter creates one page for the front cover with a list of stories, and one page per story. It also includes a navigation panel for all pages.</P>" + "<P>The layout of the newsletter is controlled with a Newsletter Template. Use HTML and the addons 'Newsletter-body only' and Newsletter-nav only' to design your look and feel.</P>" + "<P>If you will be creating an email from this newsletter, be sure to include your styles in either the newsletter template or the newsletter record.</P>" + "<P>When you view the newsletter addon for the first time, it will automatically create a 'Default' newsletter for you.</P>" + "<P>To create a new issue for this newsletter, click the 'Add a new Issue' link. The new issue will automatically appear to the publish on the publish date you set. Before the publish date only administrators can access the new issue as they add or modify stories.</P>" + "<P>To create a new newsletter, click the 'Add a new Newsletter' link. To make your new newsletter appear here, turn on Advanced Edit and click the Options icon at the top of add-on (wrench icon). Select the newsletter you want to display and hit update.</P>" + "";






                        if (!string.IsNullOrEmpty(Controls)) {
                            returnHtml = returnHtml + NewsletterController.GetAdminHintWrapper(CP, Controls);
                        }

                    }
                    // 
                    // Add any user errors
                    // 
                    if (!CP.UserError.OK()) {
                        returnHtml = "<div style=\"padding:10px\">" + CP.UserError.GetList() + "</div>" + returnHtml;
                    }
                    // returnHtml = GetContent(CP, refreshQueryString)
                }
            } catch (Exception ex) {
                HandleError(CP, ex, "execute");
            }
            CP.Addon.ExecuteAsProcessByUniqueName("RSS Feed Process");
            return returnHtml;
        }
        // 
        // =====================================================================================
        // common report for this class
        private void HandleError(CPBaseClass cp, Exception ex, string @method) {
            try {
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterPageClass." + @method);
            } catch (Exception) {
                //
                // stop anything thrown from cp errorReport
                //
            }
        }
        //
        //
        //
        private int CreateEmailGetID(CPBaseClass cp, int IssueID, int NewsletterID, string refreshQueryString, int currentIssueId) {
            int returnId = 0;
            try {
                string NewsletterName;
                string EmailAddress;
                string MemberName;
                var CSPointer = cp.CSNew();
                var cs = cp.CSNew();
                string templateCopy = "";
                var cn = new NewsletterController();
                NewsletterBodyClass Body;
                var webTemplateID = default(int);
                NewsletterNavClass Nav;
                string Styles;
                var layout = new BlockClass();
                string itemList = "";
                string newsNav = "";
                string emailBody = "";
                int LoopPtr;
                int StartPos;
                int EndPos;
                string newsCoverStoryItem = "";
                string itemLayoutAdBanners = "";
                string newsCoverCategoryItem = "";
                int emailTemplateID = 0;
                int templateId = 0;
                string adBannerLink;
                string mastheadFilename = "";
                string footerFilename = "";
                // 
                if (IssueID > 0) {
                    Constants.openRecord(cp, ref cs, "Newsletters", NewsletterID);
                    if (cs.OK()) {
                        webTemplateID = cs.GetInteger("TemplateID");
                        emailTemplateID = cs.GetInteger("emailTemplateID");
                        Styles = cp.CdnFiles.Read(cs.GetText("StylesFileName"));
                        mastheadFilename = cs.GetText("mastheadFilename");
                        footerFilename = cs.GetText("footerFilename");
                    }
                    cs.Close();
                    // 
                    templateId = emailTemplateID;
                    if (templateId != 0) {
                        // 
                        // verify it
                        Constants.openRecord(cp, ref cs, "newsletter templates", templateId);
                        if (cs.OK()) {
                            templateCopy = cs.GetText("Template");
                        }
                        cs.Close();
                    }
                    // 
                    if (string.IsNullOrEmpty(templateCopy)) {
                        // 
                        // -- no email template available, rebuild from installation file
                        templateId = NewsletterController.verifyDefaultEmailTemplateGetId(cp);
                        cp.Db.ExecuteNonQuery("update newsletters set emailTemplateID=" + templateId + " where id=" + NewsletterID);
                        Constants.openRecord(cp, ref cs, "newsletter templates", templateId);
                        if (cs.OK()) {
                            templateCopy = cs.GetText("Template");
                        }
                        cs.Close();
                    }
                    // 
                    if (string.IsNullOrEmpty(templateCopy)) {
                        // 
                        // -- if all else fails, use web template
                        Constants.openRecord(cp, ref cs, "Newsletter Templates", webTemplateID);
                        if (cs.OK()) {
                            templateCopy = cs.GetText("Template");
                        }
                        cs.Close();
                    }
                }
                // 
                // There is a template, encoding it captures the newsletterBodyClass
                // 
                string sponsor = "";
                var publishDate = DateTime.MinValue;
                string tagLine = "";
                string emailLinkToWebHtml;
                string qs;
                // 
                layout.load(templateCopy);
                if (!string.IsNullOrEmpty(mastheadFilename)) {
                    mastheadFilename = Uri.EscapeUriString(mastheadFilename);
                    layout.setClassInner("newsHeaderMasthead", "<img width=\"100%\" src=\"" + cp.Http.CdnFilePathPrefix + mastheadFilename + "\" class=\"banner\" />");
                }
                if (!string.IsNullOrEmpty(footerFilename)) {
                    footerFilename = Uri.EscapeUriString(footerFilename);
                    layout.setClassInner("newsFooterMasthead", "<img width=\"100%\" src=\"" + cp.Http.CdnFilePathPrefix + footerFilename + "\" class=\"footer\" />");
                }
                // 
                // set the link back to the web version
                // 
                emailLinkToWebHtml = layout.getClassInner("emailLinkToWeb");
                if (!string.IsNullOrEmpty(emailLinkToWebHtml)) {
                    qs = cp.Doc.RefreshQueryString;
                    qs = cp.Utils.ModifyQueryString(qs, "issueId", IssueID.ToString());
                    emailLinkToWebHtml = emailLinkToWebHtml.Replace("href=\"#\"", "href=\"?" + qs + "\"");
                    layout.setClassInner("emailLinkToWeb", emailLinkToWebHtml);
                }
                // 
                newsCoverStoryItem = layout.getClassOuter("newsCoverStoryItem");
                itemLayoutAdBanners = layout.getClassOuter("adBannerItem");
                newsCoverCategoryItem = layout.getClassOuter("newsCoverCategoryItem");
                Body = new NewsletterBodyClass();
                itemList = Body.GetCoverContent(cp, IssueID, 0, refreshQueryString, Constants.FormCover, newsCoverStoryItem, newsCoverCategoryItem, false, ref sponsor, ref publishDate, ref tagLine);
                // 
                // Call cp.Utils.AppendLogFile("createEmailGetId, 300")
                // 
                // '
                // ' add footer ad banner(s)
                // '
                string footerAdBanners = "";
                // 
                if (cs.Open("newsletter Issues", "id=" + IssueID)) {
                    string adBanner;
                    // Dim adBannerLink As String
                    int bannerLayoutId;
                    int adBannerRowCnt = 1;
                    int adBannerColumnCnt = 1;
                    int pxColumnSpace = 0;
                    int pxRowSpace = 0;
                    // 
                    bannerLayoutId = cs.GetInteger("bannerLayoutId");
                    if (bannerLayoutId > 0) {
                        var csLayout = cp.CSNew();
                        if (csLayout.Open("Newsletter Ad Banner Layouts", "id=" + bannerLayoutId)) {
                            adBannerRowCnt = csLayout.GetInteger("rowCnt");
                            adBannerColumnCnt = csLayout.GetInteger("columnCnt");
                            pxColumnSpace = csLayout.GetInteger("pxColumnSpace");
                            pxRowSpace = csLayout.GetInteger("pxRowSpace");
                        }
                        csLayout.Close();
                    }

                    for (int rowPtr = 0, loopTo = adBannerRowCnt - 1; rowPtr <= loopTo; rowPtr++) {
                        if (pxRowSpace > 0 & rowPtr > 0) {
                            footerAdBanners += @"<img src=""\cclib\images\spacer.gif"" width=""10"" height=""" + pxRowSpace.ToString() + "\" style=\"height:" + pxRowSpace.ToString() + "px\">";
                        }
                        footerAdBanners += "<div class=\"newsletterAdvertisementRow\">";
                        for (int columnPtr = 0, loopTo1 = adBannerColumnCnt - 1; columnPtr <= loopTo1; columnPtr++) {
                            if (pxColumnSpace > 0 & columnPtr > 0) {
                                footerAdBanners += @"<img src=""\cclib\images\spacer.gif"" width=""" + pxColumnSpace.ToString() + "\" height=\"10\" style=\"width:" + pxColumnSpace.ToString() + "px\">";
                            }
                            int adPtr = rowPtr * adBannerColumnCnt + columnPtr;
                            adBanner = cs.GetText("adBanner" + adPtr);
                            if (!string.IsNullOrEmpty(adBanner)) {
                                adBannerLink = cs.GetText("adBannerLink" + adPtr);
                                if (string.IsNullOrEmpty(adBannerLink)) {
                                    adBanner = Uri.EscapeUriString(adBanner);
                                    footerAdBanners += "<img src=\"" + cp.Http.CdnFilePathPrefix + adBanner + "\">";
                                } else {
                                    if (adBannerLink.IndexOf("://") < 0) {
                                        adBannerLink = "http://" + adBannerLink;
                                    }
                                    adBanner = Uri.EscapeUriString(adBanner);
                                    adBannerLink = Uri.EscapeUriString(adBannerLink);
                                    footerAdBanners += "<a href=\"" + adBannerLink + "\" target=\"_blank\"><img src=\"" + cp.Http.CdnFilePathPrefix + adBanner + "\"></a>";
                                }
                            }
                        }
                        footerAdBanners += "</div>";
                    }
                }
                cs.Close();
                if (!string.IsNullOrEmpty(footerAdBanners)) {
                    var adBannerLayout = new BlockClass();
                    adBannerLayout.load(itemLayoutAdBanners);
                    adBannerLayout.setClassInner("newsletterAdvertisements", footerAdBanners);
                    itemList += adBannerLayout.getHtml();
                }
                // 
                newsNav = layout.getClassInner("newsNav");
                Nav = new NewsletterNavClass();
                newsNav = Nav.GetNav(cp, IssueID, NewsletterID, false, 0, newsNav, currentIssueId);
                // 
                layout.setClassInner("newsNav", newsNav);
                layout.setClassInner("newsCoverList", itemList);
                layout.setClassOuter("newsBody", "");
                layout.setClassOuter("newsArchive", "");
                layout.setClassOuter("newsSearch", "");
                layout.setClassInner("newsIssueCaption", cp.Content.GetRecordName(Constants.ContentNameNewsletterIssues, IssueID));
                layout.setClassInner("newsIssueSponsor", sponsor);
                layout.setClassInner("newsIssuePublishDate", publishDate.ToShortDateString());
                if (string.IsNullOrEmpty(tagLine)) {
                    // 
                    layout.setClassOuter("newsletterTagLineRow", "");
                } else {
                    // 
                    layout.setClassInner("newsletterTagLine", tagLine);
                }
                // 
                // Add archive link
                // 
                string newsArchiveLink = layout.getClassInner("newsArchiveLink");
                newsArchiveLink = newsArchiveLink.Replace("#", cp.Utils.ModifyLinkQueryString("?" + refreshQueryString, "formId", Constants.FormArchive.ToString()));
                layout.setClassInner("newsArchiveLink", newsArchiveLink);
                // 
                emailBody = layout.getHtml();
                // 
                // Remove comments - dont know why, but emails fail with comments embedded
                // 
                LoopPtr = 0;
                while (emailBody.IndexOf("<!--", StringComparison.Ordinal) >= 0 & LoopPtr < 100) {
                    StartPos = emailBody.IndexOf("<!--", StringComparison.Ordinal);
                    EndPos = emailBody.IndexOf("-->", StartPos, StringComparison.Ordinal);
                    if (EndPos >= 0) {
                        emailBody = emailBody.Substring(0, StartPos) + emailBody.Substring(EndPos + 3);
                    }
                    LoopPtr = LoopPtr + 1;
                }
                // 
                cs.Insert(Constants.ContentNameGroupEmail);
                if (cs.OK()) {
                    returnId = cs.GetInteger("ID");
                    NewsletterName = cp.Content.GetRecordName(Constants.ContentNameNewsletterIssues, IssueID);
                    EmailAddress = (cp.User.Email ?? "").Trim();
                    MemberName = cp.User.Name;
                    if (!string.IsNullOrEmpty(EmailAddress) & !string.IsNullOrEmpty(MemberName)) {
                        EmailAddress = "\"" + MemberName + "\" <" + EmailAddress + ">";
                    }
                    cs.SetField("Name", "Newsletter " + NewsletterName);
                    cs.SetField("Subject", NewsletterName);
                    cs.SetField("FromAddress", EmailAddress);
                    cs.SetField("TestMemberID", cp.User.Id.ToString());
                    cs.SetField("CopyFileName", emailBody);
                    cs.Save();
                }
                cs.Close();
                // 
                // Call cp.Utils.AppendLogFile("createEmailGetId, 999")
                // 
            } catch (Exception ex) {
                HandleError(cp, ex, "CreateEmailGetID");
            }
            return returnId;
        }

    }
}