using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Views {
    // 
    // ====================================================================================================
    /// <summary>
    /// Design block with a centered headline, image, paragraph text and a button.
    /// </summary>
    public class NewsletterOnInstallClass : AddonBaseClass {
        // 
        // ====================================================================================================
        // 
        public override object Execute(CPBaseClass CP) {
            try {
                // 
                // -- read instanceId, guid created uniquely for this instance of the addon on a page
                int versionLastInstalled = CP.Site.GetInteger(Constants.spName_NewsletterVersion, 0);
                if (Constants.newsletterVersion > versionLastInstalled) {
                    // 
                    // -- upgraded needed
                    if (versionLastInstalled <= 3) {
                        // 
                        // -- remove old settings addon (needed until deprecated property is included in metadata)
                        CP.Db.ExecuteNonQuery("delete from ccaggregatefunctions where ccguid='{fa787411-f505-433d-990b-47bb55473ef0}'");
                    }

                }
                // 
                // -- update layout
                CP.Layout.updateLayout(Constants.guidLayoutDefaultTemplate, "Newsletter Template Default", "NewsletterTemplate.html");
                CP.Layout.updateLayout(Constants.guidLayoutDefaultIssueCover, "Newsletter Issue Default", "Newsletter Issue Default.html");
                CP.Layout.updateLayout(Constants.guidLayoutDefaultStoryOverview, "Newsletter Default Story Overview", "Newsletter Default Story Overview.html");
                CP.Layout.updateLayout(Constants.guidLayoutDefaultStoryBody, "Newsletter Default Story Copy", "Newsletter Default Story Copy.html");
                CP.Layout.updateLayout(Constants.guidLayoutDefaultEmailTemplate, "Newsletter Template Default Email", "Newsletter Template Default Email.html");
                //
                return string.Empty;
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                return "<!-- Unexpected Exception -->";
            }
        }
    }
}