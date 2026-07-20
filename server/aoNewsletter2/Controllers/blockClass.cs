

namespace Contensive.Addons.Newsletter.Controllers {
    /// <summary>
    /// A replacement for cp.block using htmlAgility
    /// </summary>
    public class BlockClass {
        // 
        private HtmlAgilityPack.HtmlDocument layout2;
        // 
        public void load(string html) {
            layout2 = new HtmlAgilityPack.HtmlDocument();
            layout2.LoadHtml(html);
        }
        // 
        public void prepend(string html) {
            layout2.Load(html + getHtml());
        }
        // 
        public void append(string html) {
            layout2.Load(getHtml() + html);
        }
        // 
        public string getClassInner(string findClass) {
            var node = layout2.DocumentNode.SelectSingleNode($"//*[contains(concat(' ', normalize-space(@class), ' '), ' {findClass} ')]");
            if (node is not null) {
                return node.InnerHtml;
            }
            return string.Empty;
        }
        // 
        public string getClassOuter(string findClass) {
            var node = layout2.DocumentNode.SelectSingleNode($"//*[contains(concat(' ', normalize-space(@class), ' '), ' {findClass} ')]");
            if (node is not null) {
                return node.OuterHtml;
            }
            return string.Empty;
        }
        // 
        public void setClassInner(string findClass, string replacement) {
            var nodes = layout2.DocumentNode.SelectNodes($"//*[contains(concat(' ', normalize-space(@class), ' '), ' {findClass} ')]");
            if (nodes is not null) {
                foreach (HtmlAgilityPack.HtmlNode node in nodes)
                    node.InnerHtml = replacement;
            }
        }
        // 
        public void setClassOuter(string findClass, string replacement) {
            var nodes = layout2.DocumentNode.SelectNodes($"//*[contains(concat(' ', normalize-space(@class), ' '), ' {findClass} ')]");
            if (nodes is not null) {
                foreach (HtmlAgilityPack.HtmlNode node in nodes) {
                    var newNode = HtmlAgilityPack.HtmlNode.CreateNode(replacement);
                    node.ParentNode.ReplaceChild(newNode, node);
                }
            }
        }
        // 
        public string getHtml() {
            if (layout2 is not null) {
                if (layout2.DocumentNode is not null) {
                    return layout2.DocumentNode.OuterHtml;
                }
            }
            return string.Empty;
        }
    }
}