using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter {

    public class RssFeedController {
        // 
        public static void updateRSSFeed(CPBaseClass cp) {
            // 
            cp.Db.ExecuteNonQuery($"update ccaggregatefunctions set processRunOnce=1 where name='RSS Feed Process'");
            // 
        }
    }
}