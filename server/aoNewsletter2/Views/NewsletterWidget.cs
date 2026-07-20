using System;
using Contensive.Addons.Newsletter.Controllers;

using Contensive.Addons.Newsletter.Models.Db;
using Contensive.Addons.Newsletter.Models.View;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Views {
    // 
    // ====================================================================================================
    /// <summary>
    /// Design block with a centered headline, image, paragraph text and a button.
    /// </summary>
    public class NewsletterWidget : AddonBaseClass {
        // 
        // ====================================================================================================
        // 
        public override object Execute(CPBaseClass CP) {
            const string designBlockName = "Newsletter Widget";
            try {
                // 
                // -- read instanceId, guid created uniquely for this instance of the addon on a page
                string result = string.Empty;
                string settingsGuid = DesignBlockController.getSettingsGuid(CP, designBlockName, ref result);
                if (string.IsNullOrEmpty(settingsGuid))
                    return result;
                // 
                // -- locate or create a data record for this guid
                int newsletterId = NewsletterController.getNewsletterId(CP, settingsGuid);
                int currentIssueID = NewsletterController.GetCurrentIssueID(CP, newsletterId);
                var settings = NewsletterModel.createOrAddSettings(CP, settingsGuid);
                if (settings is null)
                    throw new ApplicationException("Could not create the design block settings record.");
                // 
                // -- create legacy newsletter
                // 
                string legacyNewsletter = new NewsletterClass().getLegacyNewsletter(CP, newsletterId, currentIssueID)?.ToString() ?? "";
                // 
                // -- translate the Db model to a view model and mustache it into the layout
                var viewModel = NewsletterViewModel.create(CP, settings, legacyNewsletter);
                if (viewModel is null)
                    throw new ApplicationException("Could not create design block view model.");
                result = Nustache.Core.Render.StringToString(My.Resources.Resources.NewsletterLayout, viewModel);
                // 
                // -- if editing enabled, add the link and wrapperwrapper
                return CP.Content.GetEditWrapper(result, NewsletterModel.tableMetadata.contentName, settings.id);
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex);
                return "<!-- " + designBlockName + ", Unexpected Exception -->";
            }
        }
    }
}