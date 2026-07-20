using Contensive.Models.Db;

namespace Contensive.Addons.Newsletter.Models.Db {
    public class NewsletterAdBannerLayoutModel : DbBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; private set; } = new DbBaseTableMetadataModel("Newsletter Ad Banner Layouts", "NewsletterAdBannerLayouts", "default", false);
        // 
        public int rowcnt { get; set; }
        public int columncnt { get; set; }
        public int pxcolumnspace { get; set; }
        public int pxrowspace { get; set; }
        // 
    }
}