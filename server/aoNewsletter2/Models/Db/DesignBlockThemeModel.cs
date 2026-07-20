
using Contensive.Models.Db;

namespace Contensive.Addons.Newsletter.Models {
    public class DesignBlockThemeModel : DbBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; private set; } = new DbBaseTableMetadataModel("Design Block Themes", "dbThemes", "default", false);

    }
}