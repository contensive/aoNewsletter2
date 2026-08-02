using System;
using Contensive.Addons.Newsletter.Views;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that displays editable detail for a single newsletter issue.
    /// Shows Name, Publish Date, and Cover Copy fields.
    /// </summary>
    public class NewsletterIssueDetailAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterIssueDetail;
        public const string guidAddon = Constants.guidAddonNewsletterIssueDetail;
        //
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return "<p>You are not authorized to access this feature.</p>";
                }
                if (!cp.AdminUI.EndpointContainsPortal()) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                int issueId = cp.Doc.GetInteger(Constants.rnIssueId);
                int newsletterId = cp.Doc.GetInteger(Constants.rnNewsletterId);
                if (issueId == 0) {
                    if (newsletterId != 0) {
                        return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueList, $"{Constants.rnNewsletterId}={newsletterId}");
                    }
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                if (newsletterId == 0) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}", "", false, "newsletterid");
                        if (cs.OK()) {
                            newsletterId = cs.GetInteger("newsletterid");
                        }
                        cs.Close();
                    }
                }
                processForm(cp, issueId, newsletterId);
                return getForm(cp, issueId, newsletterId);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static void processForm(CPBaseClass cp, int issueId, int newsletterId) {
            try {
                if (!cp.Doc.IsProperty(Constants.rnButton)) { return; }
                string button = cp.Doc.GetText(Constants.rnButton);
                if (button == Constants.buttonSave || button == Constants.buttonOK) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}");
                        if (cs.OK()) {
                            cs.SetFormInput("name", "rnIssueName");
                            cs.SetFormInput("publishdate", "rnIssuePublishDate");
                            cs.SetFormInput("cover", "rnIssueCover");
                        }
                        cs.Close();
                    }
                }
                if (button == Constants.buttonEmailVersion) {
                    string refreshQueryString = cp.Doc.RefreshQueryString;
                    int emailId = NewsletterClass.CreateEmailGetID(cp, issueId, newsletterId, refreshQueryString, issueId);
                    if (emailId > 0) {
                        cp.Response.Redirect($"{cp.Site.GetText("adminUrl")}?cid={cp.Content.GetID(Constants.ContentNameGroupEmail)}&id={emailId}&af=4");
                    }
                    return;
                }
                if (button == Constants.buttonDelete) {
                    if (issueId > 0) {
                        cp.Content.Delete(Constants.ContentNameNewsletterIssues, $"id={issueId}");
                    }
                    string issueListLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueList)
                        + $"&{Constants.rnNewsletterId}={newsletterId}";
                    cp.Response.Redirect(issueListLink);
                    return;
                }
                if (button == Constants.buttonCancel || button == Constants.buttonOK) {
                    string issueListLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueList) + $"&{Constants.rnNewsletterId}={newsletterId}";
                    cp.Response.Redirect(issueListLink);
                    return;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static string getForm(CPBaseClass cp, int issueId, int newsletterId) {
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
                var layoutBuilder = cp.AdminUI.CreateLayoutBuilderNameValue();
                layoutBuilder.portalSubNavTitle = $"{newsletterName}, #{newsletterId}";
                //
                using (var cs = cp.CSNew()) {
                    cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}");
                    if (!cs.OK()) {
                        layoutBuilder.title = "Unknown Issue";
                        layoutBuilder.addRow();
                        layoutBuilder.rowName = "&nbsp;";
                        layoutBuilder.rowValue = $"Issue [{issueId}] was not found.";
                        return layoutBuilder.getHtml();
                    }
                    //
                    string issueName = cs.GetText("name");
                    layoutBuilder.title = string.IsNullOrWhiteSpace(issueName) ? "Issue Detail" : $"Issue: {issueName}";
                    //
                    // -- Name
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Name";
                    layoutBuilder.rowValue = cp.Html5.InputText("rnIssueName", 255, cs.GetText("name"), "form-control");
                    layoutBuilder.rowHelp = "The name for this issue.";
                    //
                    // -- Publish Date
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Publish Date";
                    DateTime publishDate = cs.GetDate("publishdate");
                    layoutBuilder.rowValue = "<div style=\"display:inline-block;width:400px\">"
                        + cp.Html5.InputDate("rnIssuePublishDate", publishDate.Equals(DateTime.MinValue) ? DateTime.MinValue : publishDate.Date, "form-control")
                        + "</div>";
                    layoutBuilder.rowHelp = "The date this issue is published. Issues are not visible before this date.";
                    //
                    // -- Cover Copy
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Cover Copy";
                    layoutBuilder.rowValue = cp.Html.InputWysiwyg("rnIssueCover", cs.GetText("cover"), CPHtmlBaseClass.EditorUserScope.Administrator);
                    layoutBuilder.rowHelp = "The cover copy displayed at the top of the newsletter main page.";
                    //
                    cs.Close();
                }
                //
                // -- layout settings
                layoutBuilder.includeForm = true;
                layoutBuilder.includeBodyColor = true;
                layoutBuilder.includeBodyPadding = true;
                layoutBuilder.isOuterContainer = false;
                //
                // -- buttons
                layoutBuilder.addFormButton(Constants.buttonCancel);
                layoutBuilder.addFormButton(Constants.buttonSave);
                layoutBuilder.addFormButton(Constants.buttonOK);
                layoutBuilder.addFormButton(Constants.buttonEmailVersion);
                layoutBuilder.addFormButton(Constants.buttonDelete, Constants.rnButton, "", "btn btn-danger mr-auto");
                //
                // -- hiddens
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterIssueDetail);
                layoutBuilder.addFormHidden(Constants.rnIssueId, issueId);
                layoutBuilder.addFormHidden(Constants.rnNewsletterId, newsletterId);
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnIssueId, issueId);
                cp.Doc.AddRefreshQueryString(Constants.rnNewsletterId, newsletterId);
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterIssueDetail);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
