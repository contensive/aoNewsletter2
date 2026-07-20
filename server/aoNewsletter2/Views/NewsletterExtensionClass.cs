
using System;
using Contensive.Addons.Newsletter.Controllers;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Views {
    // 
    // Sample Vb addon
    // 
    public class NewsletterExtensionClass : AddonBaseClass {
        // 
        // - update references to your installed version of cpBase
        // - Verify project root name space is empty
        // - Change the namespace to the collection name
        // - Change this class name to the addon name
        // - Create a Contensive Addon record with the namespace apCollectionName.ad
        // - add reference to CPBase.DLL, typically installed in c:\program files\kma\contensive\
        // 
        // =====================================================================================
        // 
        // =====================================================================================
        // 
        public override object Execute(CPBaseClass CP) {
            string returnHtml = "";
            try {
                returnHtml = "Visual Studio Contensive Addon - OK response";
            } catch (Exception ex) {
                handleError(CP, ex, "execute");
            }
            return returnHtml;
        }
        // 
        // =====================================================================================
        // common report for this class
        // =====================================================================================
        // 
        private void handleError(CPBaseClass cp, Exception ex, string @method) {
            try {
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterExtensionClass." + @method);
            } catch (Exception) {
                //
                // stop anything thrown from cp errorReport
                //
            }
        }
        // 
        // ===========================================================================================================
        // Tag Type
        // Issue - The copy is the same for each issue of the newsletter
        // Page - The copy is new for each page
        // ===========================================================================================================
        // 
        public string GetContent(CPBaseClass cp, string OptionString) {
            string returnHtml = "";
            // 
            string ExtensionName;
            string ExtensionType;
            int PageID;
            int IssueID;
            bool IsQuickEditing;
            var cn = new NewsletterController();
            string NewsletterProperty;
            string[] Parts;
            var NewsletterID = default(int);
            int currentIssueId;
            // 
            if (true) {
                // 
                // Assume newsletterNavClass is used within a PageClass
                // Get the Issue and Newsletter from the visit properties set in PageClass
                // 
                ExtensionName = (cp.Doc.GetText("ExtensionName", OptionString) ?? "").Trim();
                NewsletterProperty = cp.Visit.GetText(Constants.VisitPropertyNewsletter);
                Parts = NewsletterProperty.Split(new[] { "." }, StringSplitOptions.None);
                if (Parts.Length - 1 > 2) {
                    NewsletterID = cp.Utils.EncodeInteger(Parts[0]);
                    IssueID = cp.Utils.EncodeInteger(Parts[1]);
                    PageID = cp.Utils.EncodeInteger(Parts[2]);
                    // FormID = cp.Utils.EncodeInteger(Parts(3))
                }
                if (string.IsNullOrEmpty(ExtensionName)) {
                    ExtensionName = "Default";
                    // If Main.IsAdmin() Then
                    // returnHtml = NewsletterController.getAdminHintWrapper( cp,"The ExtensionName is blank. To use the Page Extension, set the ExtensionName and select the ExtensionType.")
                    // End If
                } else {
                    // 
                    // Handle PageID Request Variable
                    // 
                    currentIssueId = NewsletterController.GetCurrentIssueID(cp, NewsletterID);
                    ExtensionType = (cp.Doc.GetText("ExtensionType", OptionString) ?? "").Trim().ToLowerInvariant();
                    cp.Site.TestPoint("GetIssueID call 1, NewsletterID=" + NewsletterID);
                    IssueID = NewsletterController.GetIssueID(cp, NewsletterID, currentIssueId);
                    PageID = cp.Doc.GetInteger(Constants.RequestNameStoryId);
                    IsQuickEditing = cp.User.IsQuickEditing("Page Content");
                    // 
                    switch (ExtensionType ?? "") {
                        case "issue": {
                                if (IssueID != 0) {
                                    returnHtml = cp.Content.GetCopy("Newsletter-Extension-Issue-" + IssueID + "-" + ExtensionName);
                                }

                                break;
                            }
                        case "page": {
                                if (PageID != 0) {
                                    returnHtml = cp.Content.GetCopy("Newsletter-Extension-Issue-Page-" + IssueID + "-" + PageID + "-" + ExtensionName);
                                }

                                break;
                            }

                        default: {
                                returnHtml = NewsletterController.GetAdminHintWrapper(cp, "The Extension Type is blank. To use the Page Extension, set the ExtensionName and select the ExtensionType.");
                                break;
                            }
                    }
                }
            }
            return returnHtml;
        }

    }
}