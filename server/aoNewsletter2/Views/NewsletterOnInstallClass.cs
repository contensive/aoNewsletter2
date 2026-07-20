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
                CP.Layout.updateLayout("{61913e8e-bf87-4b35-8fbf-c4b4ec4f1219}", "Newsletter Template - 2025", @"newsletterAddon\NewsletterTemplate.html");
                // 
                return string.Empty;
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                return "<!-- Unexpected Exception -->";
            }
        }
    }
}