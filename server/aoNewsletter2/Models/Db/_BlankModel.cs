using Contensive.Models.Db;

namespace Contensive.Addons.Newsletter.Models.Db {
    public class BlankModel : DbBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; private set; } = new DbBaseTableMetadataModel("Newsletter Ad Banner Layouts", "NewsletterAdBannerLayouts", "default", false);
        // 
        public string imageFilename { get; set; }
        // 
    }
}