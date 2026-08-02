using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that displays editable ad banner settings for a newsletter issue.
    /// Shows Banner Campaign, Ad Banner Layout, and 6 banner image/link pairs.
    /// </summary>
    public class NewsletterIssueAdBannersAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterIssueAdBanners;
        public const string guidAddon = Constants.guidAddonNewsletterIssueAdBanners;
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
                            cs.SetFormInput("bannercampaignid", "rnBannerCampaignId");
                            cs.SetFormInput("bannerlayoutid", "rnBannerLayoutId");
                            cs.SetFormInput("adbanner0", "rnAdBanner0");
                            cs.SetFormInput("adbannerlink0", "rnAdBannerLink0");
                            cs.SetFormInput("adbanner1", "rnAdBanner1");
                            cs.SetFormInput("adbannerlink1", "rnAdBannerLink1");
                            cs.SetFormInput("adbanner2", "rnAdBanner2");
                            cs.SetFormInput("adbannerlink2", "rnAdBannerLink2");
                            cs.SetFormInput("adbanner3", "rnAdBanner3");
                            cs.SetFormInput("adbannerlink3", "rnAdBannerLink3");
                            cs.SetFormInput("adbanner4", "rnAdBanner4");
                            cs.SetFormInput("adbannerlink4", "rnAdBannerLink4");
                            cs.SetFormInput("adbanner5", "rnAdBanner5");
                            cs.SetFormInput("adbannerlink5", "rnAdBannerLink5");
                        }
                        cs.Close();
                    }
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
                    layoutBuilder.title = string.IsNullOrWhiteSpace(issueName) ? "Ad Banners" : $"Ad Banners: {issueName}";
                    //
                    // -- Banner Campaign
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Banner Campaign";
                    layoutBuilder.rowValue = cp.Html.SelectContent("rnBannerCampaignId", cs.GetInteger("bannercampaignid").ToString(), "Banner Campaigns", "", "form-control");
                    layoutBuilder.rowHelp = "Select a banner campaign for this issue.";
                    //
                    // -- Ad Banner Layout
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Ad Banner Layout";
                    layoutBuilder.rowValue = cp.Html.SelectContent("rnBannerLayoutId", cs.GetInteger("bannerlayoutid").ToString(), "Newsletter Ad Banner Layouts", "", "form-control");
                    layoutBuilder.rowHelp = "Select the layout for ad banner presentation.";
                    //
                    // -- Banner 1 through 6
                    for (int i = 0; i < 6; i++) {
                        int bannerNum = i + 1;
                        string currentFile = cs.GetText($"adbanner{i}");
                        string currentLink = cs.GetText($"adbannerlink{i}");
                        //
                        layoutBuilder.addRow();
                        layoutBuilder.rowName = $"Banner {bannerNum} Image";
                        string fileHtml = "";
                        if (!string.IsNullOrWhiteSpace(currentFile)) {
                            fileHtml = $"<div class=\"mb-2\">Current: {currentFile}</div>";
                        }
                        fileHtml += cp.Html.InputFile($"rnAdBanner{i}");
                        layoutBuilder.rowValue = fileHtml;
                        layoutBuilder.rowHelp = $"Upload an image for banner {bannerNum}.";
                        //
                        layoutBuilder.addRow();
                        layoutBuilder.rowName = $"Banner {bannerNum} Link";
                        layoutBuilder.rowValue = cp.Html5.InputText($"rnAdBannerLink{i}", 255, currentLink, "form-control");
                        layoutBuilder.rowHelp = $"The URL to navigate to when banner {bannerNum} is clicked.";
                    }
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
                //
                // -- hiddens
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterIssueAdBanners);
                layoutBuilder.addFormHidden(Constants.rnIssueId, issueId);
                layoutBuilder.addFormHidden(Constants.rnNewsletterId, newsletterId);
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnIssueId, issueId);
                cp.Doc.AddRefreshQueryString(Constants.rnNewsletterId, newsletterId);
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterIssueAdBanners);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
