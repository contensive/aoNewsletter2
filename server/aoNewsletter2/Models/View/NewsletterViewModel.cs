using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Models.View {
    public class NewsletterViewModel : DesignBlockViewBaseModel {
        // 
        public string legacyNewsletter { get; set; }
        // 
        // ====================================================================================================
        /// <summary>
        /// Populate the view model from the entity model
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        public static NewsletterViewModel create(CPBaseClass cp, Db.NewsletterModel settings, string legacyNewsletter) {
            try {
                // 
                // -- base fields
                var result = create<NewsletterViewModel>(cp, settings);
                // 
                // -- custom
                result.legacyNewsletter = legacyNewsletter;
                // 
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return null;
            }
        }
    }

}