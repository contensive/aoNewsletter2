using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.Addons.Newsletter.Models.Db {
    public class NewsletterModel : DesignBlockBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; private set; } = new DbBaseTableMetadataModel("Newsletters", "Newsletters", "default", false);
        // 
        public int templateId { get; set; }
        public FieldTypeCSSFile stylesFileName { get; set; }
        public int emailTemplateId { get; set; }
        public string mastheadFilename { get; set; }
        public string footerFilename { get; set; }
        public bool blockArchiveSearchForm { get; set; }
        public int archiveIssuesToDisplay { get; set; }
        public int searchResultsPerPage { get; set; }

        public static NewsletterModel createOrAddSettings(CPBaseClass cp, string settingsGuid) {
            var result = create<NewsletterModel>(cp, settingsGuid);

            if (result is null) {
                result = addDefault<NewsletterModel>(cp);
                result.name = tableMetadata.contentName + " " + result.id;
                result.ccguid = settingsGuid;
                result.themeStyleId = 0;
                result.padTop = false;
                result.padBottom = false;
                result.padRight = false;
                result.padLeft = false;

                result.save(cp);
                cp.Content.LatestContentModifiedDate.Track(result.modifiedDate);
            }

            return result;
        }
    }
}