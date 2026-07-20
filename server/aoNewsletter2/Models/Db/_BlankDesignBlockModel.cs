using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.Addons.Newsletter.Models.Db {
    public class BlankDesignBlockModel : DesignBlockBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; private set; } = new DbBaseTableMetadataModel("blank", "blank", "default", false);
        public string imageFilename { get; set; }
        public string headline { get; set; }
        public string embed { get; set; }
        public string description { get; set; }
        public string buttonText { get; set; }
        public string buttonUrl { get; set; }
        public int imageAspectRatioId { get; set; }
        public string btnStyleSelector { get; set; }

        public static BlankDesignBlockModel createOrAddSettings(CPBaseClass cp, string settingsGuid) {
            var result = create<BlankDesignBlockModel>(cp, settingsGuid);

            if (result is null) {
                result = addDefault<BlankDesignBlockModel>(cp);
                result.name = tableMetadata.contentName + " " + result.id;
                result.ccguid = settingsGuid;
                result.themeStyleId = 0;
                result.padTop = false;
                result.padBottom = false;
                result.padRight = false;
                result.padLeft = false;
                result.imageFilename = string.Empty;
                result.imageAspectRatioId = 3;
                result.headline = "Lorem Ipsum Dolor";
                result.description = "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>";
                result.embed = string.Empty;
                result.buttonUrl = string.Empty;
                result.buttonText = string.Empty;
                result.save(cp);
                cp.Content.LatestContentModifiedDate.Track(result.modifiedDate);
            }

            return result;
        }
    }
}