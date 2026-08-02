using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that displays editable detail for a single newsletter.
    /// Modeled on BlogDetailsAddon from aoBlog.
    /// </summary>
    public class NewsletterDetailAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterDetail;
        public const string guidAddon = Constants.guidAddonNewsletterDetail;
        //
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return "<p>You are not authorized to access this feature.</p>";
                }
                if (!cp.AdminUI.EndpointContainsPortal()) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                int newsletterId = cp.Doc.GetInteger(Constants.rnNewsletterId);
                if (newsletterId == 0) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                processForm(cp, newsletterId);
                return getForm(cp, newsletterId);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static void processForm(CPBaseClass cp, int newsletterId) {
            try {
                if (!cp.Doc.IsProperty(Constants.rnButton)) { return; }
                string button = cp.Doc.GetText(Constants.rnButton);
                if (button == Constants.buttonSave || button == Constants.buttonOK) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletters, $"id={newsletterId}");
                        if (cs.OK()) {
                            cs.SetFormInput("name", "rnNewsletterName");
                            cs.SetFormInput("searchresultsperpage", "rnSearchResultsPerPage");
                            cs.SetFormInput("archiveissuestodisplay", "rnArchiveIssuesToDisplay");
                            cs.SetFormInput("blockarchivesearchform", "rnBlockArchiveSearchForm");
                        }
                        cs.Close();
                    }
                }
                if (button == Constants.buttonDelete) {
                    if (newsletterId > 0) {
                        cp.Content.Delete(Constants.ContentNameNewsletters, $"id={newsletterId}");
                    }
                    cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                    return;
                }
                if (button == Constants.buttonCancel || button == Constants.buttonOK) {
                    cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                    return;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static string getForm(CPBaseClass cp, int newsletterId) {
            try {
                if (!cp.Response.isOpen) { return ""; }
                //
                var layoutBuilder = cp.AdminUI.CreateLayoutBuilderNameValue();
                layoutBuilder.title = "Newsletter Detail";
                layoutBuilder.includeForm = true;
                layoutBuilder.includeBodyColor = true;
                layoutBuilder.includeBodyPadding = true;
                layoutBuilder.isOuterContainer = false;
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterDetail);
                cp.Doc.AddRefreshQueryString(Constants.rnNewsletterId, newsletterId);
                //
                using (var cs = cp.CSNew()) {
                    cs.Open(Constants.ContentNameNewsletters, $"id={newsletterId}");
                    if (!cs.OK()) {
                        layoutBuilder.portalSubNavTitle = "Unknown Newsletter";
                        layoutBuilder.addRow();
                        layoutBuilder.rowName = "&nbsp;";
                        layoutBuilder.rowValue = $"Newsletter [{newsletterId}] was not found.";
                        return layoutBuilder.getHtml();
                    }
                    //
                    layoutBuilder.portalSubNavTitle = $"{cs.GetText("name")}, #{cs.GetInteger("id")}";
                    //
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Name";
                    layoutBuilder.rowValue = cp.Html5.InputText("rnNewsletterName", 255, cs.GetText("name"), "form-control");
                    layoutBuilder.rowHelp = "The name for this newsletter.";
                    //
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Search Results Per Page";
                    layoutBuilder.rowValue = cp.Html5.InputText("rnSearchResultsPerPage", 10, cs.GetInteger("searchresultsperpage").ToString(), "form-control");
                    layoutBuilder.rowHelp = "The number of archive issues to display in the archive issue list.";
                    //
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Archive Issues To Display";
                    layoutBuilder.rowValue = cp.Html5.InputText("rnArchiveIssuesToDisplay", 10, cs.GetInteger("archiveissuestodisplay").ToString(), "form-control");
                    layoutBuilder.rowHelp = "The number of archive issues to display.";
                    //
                    layoutBuilder.addRow();
                    layoutBuilder.rowName = "Block Archive Search Form";
                    layoutBuilder.rowValue = cp.Html5.CheckBox("rnBlockArchiveSearchForm", cs.GetBoolean("blockarchivesearchform"), "form-check-input");
                    layoutBuilder.rowHelp = "When checked, the archive search form is hidden.";
                    //
                    cs.Close();
                }
                //
                // -- buttons
                layoutBuilder.addFormButton(Constants.buttonCancel);
                layoutBuilder.addFormButton(Constants.buttonSave);
                layoutBuilder.addFormButton(Constants.buttonOK);
                layoutBuilder.addFormButton(Constants.buttonDelete, Constants.rnButton, "", "btn btn-danger mr-auto");
                //
                // -- hiddens
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterDetail);
                layoutBuilder.addFormHidden(Constants.rnNewsletterId, newsletterId);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
