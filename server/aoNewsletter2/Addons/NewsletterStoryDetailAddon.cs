using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that displays editable detail for a single newsletter story.
    /// Shows Name, Cover Overview, and Full Story body fields.
    /// Third-level navigation under Stories.
    /// </summary>
    public class NewsletterStoryDetailAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterStoryDetail;
        public const string guidAddon = Constants.guidAddonNewsletterStoryDetail;
        //
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return "<p>You are not authorized to access this feature.</p>";
                }
                if (!cp.AdminUI.EndpointContainsPortal()) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                int storyId = cp.Doc.GetInteger(Constants.rnStoryId);
                int issueId = cp.Doc.GetInteger(Constants.rnIssueId);
                int newsletterId = cp.Doc.GetInteger(Constants.rnNewsletterId);
                if (storyId == 0) {
                    if (issueId != 0) {
                        return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueStories, $"{Constants.rnIssueId}={issueId}&{Constants.rnNewsletterId}={newsletterId}");
                    }
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                //
                // -- look up issueId and newsletterId from the story if not provided
                if (issueId == 0) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId}", "", false, "newsletterid");
                        if (cs.OK()) {
                            issueId = cs.GetInteger("newsletterid");
                        }
                        cs.Close();
                    }
                }
                if (newsletterId == 0 && issueId != 0) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}", "", false, "newsletterid");
                        if (cs.OK()) {
                            newsletterId = cs.GetInteger("newsletterid");
                        }
                        cs.Close();
                    }
                }
                processForm(cp, storyId, issueId, newsletterId);
                return getForm(cp, storyId, issueId, newsletterId);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static void processForm(CPBaseClass cp, int storyId, int issueId, int newsletterId) {
            try {
                if (!cp.Doc.IsProperty(Constants.rnButton)) { return; }
                string button = cp.Doc.GetText(Constants.rnButton);
                if (button == Constants.buttonSave || button == Constants.buttonOK) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId}");
                        if (cs.OK()) {
                            cs.SetFormInput("name", "rnStoryName");
                            cs.SetFormInput("overview", "rnStoryOverview");
                            cs.SetFormInput("body", "rnStoryBody");
                        }
                        cs.Close();
                    }
                }
                if (button == Constants.buttonDelete) {
                    if (storyId > 0) {
                        cp.Content.Delete(Constants.ContentNameNewsletterStories, $"id={storyId}");
                    }
                    string storiesLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueStories)
                        + $"&{Constants.rnIssueId}={issueId}&{Constants.rnNewsletterId}={newsletterId}";
                    cp.Response.Redirect(storiesLink);
                    return;
                }
                if (button == Constants.buttonCancel || button == Constants.buttonOK) {
                    string storiesLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueStories)
                        + $"&{Constants.rnIssueId}={issueId}&{Constants.rnNewsletterId}={newsletterId}";
                    cp.Response.Redirect(storiesLink);
                    return;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static string getForm(CPBaseClass cp, int storyId, int issueId, int newsletterId) {
            try {
                if (!cp.Response.isOpen) { return ""; }
                //
                // -- get newsletter name for sub-nav title
                string newsletterName = "";
                using (var cs = cp.CSNew()) {
                    cs.Open(Constants.ContentNameNewsletters, $"id={newsletterId}", "", false, "name");
                    if (cs.OK()) {
                        newsletterName = cs.GetText("name");
                    }
                    cs.Close();
                }
                //
                // -- get story data
                string storyName = "";
                string storyOverview = "";
                string storyBody = "";
                using (var cs = cp.CSNew()) {
                    cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId}", "", false, "name,overview,body");
                    if (cs.OK()) {
                        storyName = cs.GetText("name");
                        storyOverview = cs.GetText("overview");
                        storyBody = cs.GetText("body");
                    }
                    cs.Close();
                }
                //
                var layoutBuilder = cp.AdminUI.CreateLayoutBuilder();
                layoutBuilder.portalSubNavTitle = $"{newsletterName}, #{newsletterId}";
                //
                // -- title field, overview, and WYSIWYG editor for story body
                string nameInput = cp.Html5.InputText("rnStoryName", 255, storyName, "form-control");
                string nameRow = $"<div class=\"mb-3\"><label class=\"form-label\"><b>Name</b></label>{nameInput}</div>";
                //
                string overviewLabel = "<div class=\"mb-3\"><label class=\"form-label\"><b>Cover Overview</b></label>"
                    + "<div class=\"form-text\">Overview displayed on the newsletter cover under the story link.</div></div>";
                string overviewEditor = cp.Html.InputWysiwyg("rnStoryOverview", storyOverview, CPHtmlBaseClass.EditorUserScope.Administrator);
                //
                string bodyLabel = "<div class=\"mb-3\"><label class=\"form-label\"><b>Full Story</b></label>"
                    + "<div class=\"form-text\">The full story content displayed on the story page.</div></div>";
                string bodyEditor = cp.Html.InputWysiwyg("rnStoryBody", storyBody, CPHtmlBaseClass.EditorUserScope.Administrator);
                //
                layoutBuilder.body = nameRow + overviewLabel + overviewEditor + bodyLabel + bodyEditor;
                //
                // -- layout settings
                layoutBuilder.title = string.IsNullOrWhiteSpace(storyName) ? "Story Detail" : $"Edit Story: {storyName}";
                layoutBuilder.callbackAddonGuid = Constants.guidAddonNewsletterStoryDetail;
                layoutBuilder.includeForm = true;
                //
                // -- buttons
                layoutBuilder.addFormButton(Constants.buttonOK);
                layoutBuilder.addFormButton(Constants.buttonSave);
                layoutBuilder.addFormButton(Constants.buttonCancel);
                layoutBuilder.addFormButton(Constants.buttonDelete);
                //
                // -- hiddens
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterStoryDetail);
                layoutBuilder.addFormHidden(Constants.rnStoryId, storyId);
                layoutBuilder.addFormHidden(Constants.rnIssueId, issueId);
                layoutBuilder.addFormHidden(Constants.rnNewsletterId, newsletterId);
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnStoryId, storyId);
                cp.Doc.AddRefreshQueryString(Constants.rnIssueId, issueId);
                cp.Doc.AddRefreshQueryString(Constants.rnNewsletterId, newsletterId);
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterStoryDetail);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
