
using System;
using Contensive.Addons.Newsletter.Models.Db;
using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.Addons.Newsletter.Controllers {
    public class NewsletterController {
        //
        private static readonly Random _random = new Random();
        //
        public const string cr = "\r\n\t";
        // 
        // =====================================================================================
        // common report for this class
        // =====================================================================================
        // 
        private static void handleError(CPBaseClass cp, Exception ex, string @method) {
            try {
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterCommonClass." + @method);
            } catch (Exception) {
                //
                // stop anything thrown from cp errorReport
                //
            }
        }
        //
        internal static int GetIssueID(CPBaseClass cp, int NewsletterID, int currentIssueId) {
            // 
            int IssueID = cp.Doc.GetInteger(Constants.RequestNameIssueID);
            // 
            if (IssueID == 0) {
                IssueID = currentIssueId;
            }
            // 
            return IssueID;
        }
        // 
        internal static int GetCurrentIssueID(CPBaseClass cp, int NewsletterID) {
            try {
                int returnId = 0;
                using (var cs = cp.CSNew()) {
                    // 
                    cs.Open(Constants.ContentNameNewsletterIssues, "active=1 and (PublishDate<=" + cp.Db.EncodeSQLDate(DateTime.Now) + ") AND (NewsletterID=" + NewsletterID + ")", "PublishDate desc, ID desc");
                    if (cs.OK()) {
                        returnId = cs.GetInteger("ID");
                    }
                    cs.Close();
                }
                // 
                if (returnId == 0) {
                    // 
                    // there are no issues of this newsletter -- create a default issue
                    return createDefaultIssueGetId(cp, NewsletterID);
                }
                return returnId;
            } catch (Exception ex) {
                handleError(cp, ex, "getCurrentIssueId");
                return 0;
            }
        }
        // 
        internal static string GetUnpublishedIssueList(CPBaseClass cp, int NewsletterID, NewsletterController cn) {
            string GetUnpublishedIssueListRet = default;
            GetUnpublishedIssueListRet = "";
            // 
            string qs = "";
            var cs = cp.CSNew();
            int ID;
            string Name;
            bool Active;
            DateTime PublishDate;
            string Copy;
            DateTime DateAdded;
            bool isContentMan;
            // 
            isContentMan = cp.User.IsContentManager("Newsletters");
            cs.Open(Constants.ContentNameNewsletterIssues, "active=1 and (newsletterid=" + NewsletterID + ")and(PublishDate is null)or(PublishDate>" + cp.Db.EncodeSQLDate(DateTime.Now) + ")", "PublishDate desc, ID desc");
            while (cs.OK()) {
                ID = cs.GetInteger("ID");
                Name = (cs.GetText("name") ?? "").Trim();
                Active = cs.GetBoolean("active");
                PublishDate = cs.GetDate("PublishDate");
                DateAdded = cs.GetDate("DateAdded");
                Copy = Name;
                if (string.IsNullOrEmpty(Copy)) {
                    Copy = "unnamed #" + ID;
                }
                if (!Active) {
                    Copy = Copy + ",inactive";
                }
                if (encodeMinDate(DateAdded) != DateTime.MinValue) {
                    Copy = Copy + ", created " + DateAdded.ToShortDateString();
                }
                if (PublishDate != DateTime.MinValue) {
                    Copy = Copy + ", publish " + PublishDate.ToShortDateString();
                }
                if (isContentMan) {
                    qs = cp.Doc.RefreshQueryString;
                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameIssueID, ID.ToString());
                    Copy = "<a href=\"?" + qs + "\">" + Copy + "</a>";
                }
                GetUnpublishedIssueListRet = GetUnpublishedIssueListRet + "<li>" + Copy + "</li>";
                cs.GoNext();
            }
            cs.Close();
            // 
            if (!string.IsNullOrEmpty(GetUnpublishedIssueListRet)) {
                if (isContentMan) {
                    qs = cp.Doc.RefreshQueryString;
                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameIssueID, "");
                    GetUnpublishedIssueListRet += "<li><a href=\"?" + qs + "\">Current Issue</a></li>";
                }
                GetUnpublishedIssueListRet = "<UL>" + GetUnpublishedIssueListRet + "</UL>";
            }

            return GetUnpublishedIssueListRet;
            // 
            // Exit Function
            // ErrorTrap:
            // Call HandleError("aoNewsletter.newsletterCommonClass", "GetUnpublishedIssueList")
        }
        // 
        internal static int getNewsletterId(CPBaseClass cp, string addonArgInstanceGuid) {
            int returnId = 0;
            try {
                int addonArgumentNewsletterId = cp.Doc.GetInteger("Newsletter");
                string criteria = "";
                if (addonArgumentNewsletterId > 0) {
                    criteria = "(id=" + addonArgumentNewsletterId + ")";
                } else if (!string.IsNullOrEmpty(addonArgInstanceGuid)) {
                    criteria = "(ccguid='" + addonArgInstanceGuid + "')";
                } else {
                    criteria = "(name='Default')";
                }
                using (var cs = cp.CSNew()) {
                    if (cs.Open(Constants.ContentNameNewsletters, criteria)) {
                        returnId = cs.GetInteger("ID");
                    } else {
                        cs.Close();
                        // 
                        // must create new newsletter
                        // 
                        cs.Insert(Constants.ContentNameNewsletters);
                        if (cs.OK()) {
                            returnId = cs.GetInteger("ID");
                            verifyAdBannerLayouts(cp);
                            int templateID = verifyDefaultTemplateGetId(cp);
                            int emailTemplateID = verifyDefaultEmailTemplateGetId(cp);
                            if (!string.IsNullOrEmpty(addonArgInstanceGuid)) {
                                // 
                                // newsletter called out by guid but not found
                                // 
                                cs.SetField("ccguid", addonArgInstanceGuid);
                                cs.SetField("Name", cp.Content.GetRecordName("page content", cp.Doc.PageId));
                            } else {
                                // 
                                // all other cases
                                // 
                                cs.SetField("Name", $"Newsletter {returnId}");
                            }
                            cs.SetField("TemplateID", templateID.ToString());
                            cs.SetField("emailTemplateID", emailTemplateID.ToString());
                        }
                        createDefaultIssueGetId(cp, returnId);
                    }
                    cs.Close();
                }
            } catch (Exception ex) {
                handleError(cp, ex, "getnewsletterId");
            }
            return returnId;
        }
        // 
        internal static int createDefaultIssueGetId(CPBaseClass cp, int newsletterId) {
            int returnId = 0;
            try {
                var cs = cp.CSNew();
                string newsletterName = cp.Content.GetRecordName(Constants.ContentNameNewsletters, newsletterId);
                // 
                // Build the first issue in the newsletter
                // 
                cs.Insert("Newsletter Issues");
                if (cs.OK()) {
                    returnId = cs.GetInteger("id");
                    cs.SetField("name", newsletterName + " Newsletter, Issue 1");
                    cs.SetField("NewsletterID", newsletterId.ToString());
                    cs.SetField("PublishDate", DateTime.Now.ToShortDateString());
                    cs.SetField("cover", cp.Layout.GetLayout(Constants.guidLayoutDefaultIssueCover));
                }
                cs.Close();
                // 
                // Build the first story
                // 
                cs.Insert("Newsletter Stories");
                if (cs.OK()) {
                    cs.SetField("name", "The First Story");
                    cs.SetField("newsletterid", returnId.ToString());
                    cs.SetField("Overview", cp.Layout.GetLayout(Constants.guidLayoutDefaultStoryOverview));
                    cs.SetField("body", cp.Layout.GetLayout(Constants.guidLayoutDefaultStoryBody));
                }
                cs.Close();
            } catch (Exception ex) {
                handleError(cp, ex, "createDefaultIssueGetId");
            }
            return returnId;
        }
        // 
        internal static void SortCategoriesByIssue(CPBaseClass cp, int IssueID) {
            var cs = cp.CSNew();
            var Pointer = cp.CSNew();
            int CategoryID;
            var Sort = default(int);
            string SQL;
            string MainSQL;
            int SortArrayPointer;
            int SortArrayCount;
            string SortOrder;
            int RuleCategoryID;
            int RuleIssueID;
            int ptr = 0;
            var SortArray = new string[3, 2];
            // 
            CategoryID = cp.Doc.GetInteger(Constants.RequestNameSortUp);
            // 
            // Check for Categories without rules, since rules decide sort order of categories, no stories show if
            // associated to a category without a rule, join fails.
            // 
            SQL = "SELECT NIP.CategoryID AS CatID, NewsletterID AS IssueID ";
            SQL = SQL + "FROM NewsletterIssuePages NIP ";
            SQL = SQL + "WHERE (NIP.CategoryID Not IN (SELECT CategoryID FROM NewsletterIssueCategoryRules WHERE NewsletterIssueID=" + cp.Db.EncodeSQLNumber(IssueID) + ")) ";
            SQL = SQL + "AND (NIP.CategoryID Is Not Null)";
            // 1/19/2009 just look for IssuePages within this issue that do not have IssueCategoryRules for this issue
            SQL = SQL + "AND (NIP.NewsletterID=" + cp.Db.EncodeSQLNumber(IssueID) + ")";
            // 
            cs.OpenSQL(SQL);
            while (cs.OK()) {
                Pointer.Insert(Constants.ContentNameIssueRules);
                if (Pointer.OK()) {
                    RuleCategoryID = cs.GetInteger("CatID");
                    RuleIssueID = cs.GetInteger("IssueID");
                    SortOrder = GetSortOrder(cp, RuleCategoryID, RuleIssueID);
                    Pointer.SetField("NewsletterIssueID", RuleIssueID.ToString());
                    Pointer.SetField("Active", "1");
                    Pointer.SetField("CategoryID", RuleCategoryID.ToString());
                    Pointer.SetField("SortOrder", SortOrder);
                }
                Pointer.GoNext();
                cs.GoNext();
            }
            cs.Close();
            // 
            if (CategoryID != 0) {
                // 
                MainSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder";
                MainSQL = MainSQL + " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR";
                MainSQL = MainSQL + " Where (NIC.ID = NIR.CategoryID)";
                MainSQL = MainSQL + " AND (NIR.NewsletterIssueID=" + IssueID + ")";
                MainSQL = MainSQL + " AND (NIC.Active<>0)";
                MainSQL = MainSQL + " AND (NIR.Active<>0)";
                MainSQL = MainSQL + " ORDER BY NIR.SortOrder";
                // 
                // b/c cp has no cp.getRows
                // 
                if (cs.OpenSQL(MainSQL)) {
                    SortArrayCount = cs.GetRowCount();
                    if (SortArrayCount > 0) {
                        SortArray = new string[3, SortArrayCount];
                        while (cs.OK()) {
                            SortArray[0, ptr] = cs.GetText("categoryId");
                            SortArray[1, ptr] = cs.GetText("sortOrder");
                            ptr += 1;
                            cs.GoNext();
                        }
                        SortArrayCount = ptr;
                        var loopTo = SortArrayCount - 1;
                        for (SortArrayPointer = 0; SortArrayPointer <= loopTo; SortArrayPointer++) {
                            if (CategoryID == cp.Utils.EncodeInteger(SortArray[0, SortArrayPointer]) & SortArrayPointer != 0) {
                                SortArray[1, SortArrayPointer - 1] = PadValue(cp, Sort, 4);
                                SortArray[1, SortArrayPointer] = PadValue(cp, Sort - 10, 4);
                            } else {
                                SortArray[1, SortArrayPointer] = PadValue(cp, Sort, 4);
                            }
                            Sort = Sort + 10;
                        }
                        SortArrayPointer = 0;
                        var loopTo1 = SortArrayCount - 1;
                        for (SortArrayPointer = 0; SortArrayPointer <= loopTo1; SortArrayPointer++) {
                            SQL = "Update NewsletterIssueCategoryRules SET SortOrder=" + SortArray[1, SortArrayPointer] + " WHERE (CategoryID=" + SortArray[0, SortArrayPointer] + ") AND (NewsletterIssueID=" + cp.Db.EncodeSQLNumber(IssueID) + ")";
                            cp.Db.ExecuteNonQuery(SQL);
                        }
                    }
                }
                // 
            }
        }
        // 
        internal static string GetCategoryAccessString(CPBaseClass cp, int CategoryID) {
            string GetCategoryAccessStringRet = default;
            var cs = cp.CSNew();
            string SQL;
            string Stream = "";
            // 
            SQL = "SELECT ID ";
            SQL = SQL + "From NewsletterIssuePages ";
            SQL = SQL + "WHERE (CategoryID=" + cp.Db.EncodeSQLNumber(CategoryID) + ") ";
            SQL = SQL + "AND (ID not in(Select NewsletterPageID FROM NewsletterPageGroupRules))";
            // 
            // first scheck for any unblocked story
            // 
            cs.OpenSQL(SQL);
            if (cs.OK()) {
                // 
                // no unblocked stories, look for blocked stories
                // 
                cs.Close();
                SQL = "SELECT GR.GroupID ";
                SQL = SQL + "FROM NewsletterPageGroupRules GR, NewsletterIssuePages NIP ";
                SQL = SQL + "Where (GR.NewsletterPageID = NIP.ID) ";
                SQL = SQL + "AND (NIP.CategoryID=" + cp.Db.EncodeSQLNumber(CategoryID) + ") ";
                // 
                cs.OpenSQL(SQL);
                while (cs.OK()) {
                    if (!string.IsNullOrEmpty(Stream)) {
                        Stream += ",";
                    }
                    Stream += cs.GetInteger("GroupID").ToString();
                    cs.GoNext();
                }
                cs.Close();
            }
            cs.Close();
            // 
            // If Stream <> "" Then
            // stream &=  ","
            // End If
            // 
            GetCategoryAccessStringRet = Stream;
            return GetCategoryAccessStringRet;
        }
        // 
        internal static string GetArticleAccessString(CPBaseClass cp, int StoryID) {
            string GetArticleAccessStringRet = default;
            // 
            var cs = cp.CSNew();
            string SQL;
            string Stream = "";
            // 
            SQL = "SELECT GR.GroupID ";
            SQL = SQL + "FROM NewsletterPageGroupRules GR ";
            SQL = SQL + "Where (GR.NewsletterPageID=" + cp.Db.EncodeSQLNumber(StoryID) + ")";
            // 
            cs.OpenSQL(SQL);
            while (cs.OK()) {
                if (!string.IsNullOrEmpty(Stream)) {
                    Stream += ",";
                }
                Stream += cs.GetInteger("GroupID").ToString();
                cs.GoNext();
            }
            cs.Close();
            // 
            // If Stream <> "" Then
            // stream &=  ","
            // End If
            // 
            GetArticleAccessStringRet = Stream;
            return GetArticleAccessStringRet;
        }
        // 
        internal static bool HasAccess(CPBaseClass cp, string GroupString) {
            bool HasAccessRet = default;
            // 
            string[] ListArray;
            int ListArrayCount;
            int ListArrayPointer;
            // 
            if (cp.User.IsContentManager("Newsletters")) {
                HasAccessRet = true;
            } else if (!string.IsNullOrEmpty(GroupString)) {
                if (GroupString.IndexOf(",", StringComparison.OrdinalIgnoreCase) >= 0) {
                    ListArray = GroupString.Split(new[] { "," }, StringSplitOptions.None);
                    ListArrayCount = ListArray.Length - 1;
                    var loopTo = ListArrayCount;
                    for (ListArrayPointer = 0; ListArrayPointer <= loopTo; ListArrayPointer++) {
                        if (cp.User.IsInGroup(cp.Content.GetRecordName("Groups", int.Parse(ListArray[ListArrayPointer])))) {
                            HasAccessRet = true;
                            return HasAccessRet;
                        }
                    }
                }
            } else {
                HasAccessRet = true;
            }

            return HasAccessRet;
            // 
            // Exit Function
            // ErrorTrap:
            // Call HandleError("aoNewsletter.newsletterCommonClass", "GetArticleAccessString")
        }
        // 
        private static string PadValue(CPBaseClass cp, int Value, int StringLenghth) {
            string PadValueRet = default;
            int Counter;
            int ValueLenghth;
            string InnerValue;
            // 
            InnerValue = Value.ToString();
            ValueLenghth = (InnerValue ?? "").Length;
            // 
            if (ValueLenghth < StringLenghth) {
                var loopTo = StringLenghth - 1;
                for (Counter = ValueLenghth; Counter <= loopTo; Counter++) {
                    InnerValue = "0" + InnerValue;
                }
            }
            // 
            PadValueRet = InnerValue;
            return PadValueRet;
        }
        // 
        private static string GetSortOrder(CPBaseClass cp, int CategoryID, int IssueID) {
            string GetSortOrderRet = default;
            var cs = cp.CSNew();
            string Stream = "";
            // 
            cs.Open("Newsletter Issue Category Rules", "(CategoryID=" + CategoryID + ") AND (NewsletterIssueID=" + IssueID + ")");
            if (cs.OK()) {
                Stream = cs.GetText("SortOrder");
            }
            cs.Close();
            // 
            if (string.IsNullOrEmpty(Stream)) {
                Stream = "0";
            }
            // 
            GetSortOrderRet = Stream;
            return GetSortOrderRet;
        }
        // 
        public static void verifyAdBannerLayouts(CPBaseClass cp) {
            do {
                // 
                // -- Single Ad
                var layout = DbBaseModel.createByUniqueName<NewsletterAdBannerLayoutModel>(cp, "Single Ad");
                if (layout is null) {
                    layout = DbBaseModel.addDefault<NewsletterAdBannerLayoutModel>(cp);
                    layout.name = "Single Ad";
                    layout.rowcnt = 1;
                    layout.columncnt = 1;
                    layout.pxcolumnspace = 0;
                    layout.pxrowspace = 0;
                    layout.save(cp);
                }
            }
            while (false);
            // Do
            // '
            // ' -- Double Ad, 2 Wide
            // Dim layout As NewsletterAdBannerLayoutModel = DbBaseModel.createByUniqueName(Of NewsletterAdBannerLayoutModel)(cp, "Double Ad, 2 Wide")
            // If (layout Is Nothing) Then
            // layout = DbBaseModel.addDefault(Of NewsletterAdBannerLayoutModel)(cp)
            // layout.name = "Double Ad, 2 Wide"
            // layout.rowcnt = 1
            // layout.columncnt = 2
            // layout.pxcolumnspace = 0
            // layout.pxrowspace = 0
            // layout.save(cp)
            // End If
            // Loop While False
            do {
                // 
                // -- Double Ad, 2 Stacked
                var layout = DbBaseModel.createByUniqueName<NewsletterAdBannerLayoutModel>(cp, "Double Ad, 2 Stacked");
                if (layout is null) {
                    layout = DbBaseModel.addDefault<NewsletterAdBannerLayoutModel>(cp);
                    layout.name = "Double Ad, 2 Stacked";
                    layout.rowcnt = 2;
                    layout.columncnt = 1;
                    layout.pxcolumnspace = 0;
                    layout.pxrowspace = 0;
                    layout.save(cp);
                }
            }
            while (false);
        }
        // 
        internal static int verifyDefaultTemplateGetId(CPBaseClass cp) {
            using (var cs = cp.CSNew()) {
                // 
                // -- try default template
                cs.Open("Newsletter Templates", "name='Newsletter Template Default'");
                if (!cs.OK()) {
                    cs.Close();
                    cs.Insert("Newsletter Templates");
                    if (cs.OK()) {
                        cs.SetField("name", "Newsletter Template Default");
                    }
                }
                if (cs.OK()) {
                    // 
                    // Use the default template in their Db already
                    if (string.IsNullOrEmpty((cs.GetText("Template") ?? "").Trim())) {
                        cs.SetField("Template", cp.Layout.GetLayout(Constants.guidLayoutDefaultTemplate));
                    }
                    return cs.GetInteger("ID");
                }
                cs.Close();
                return 0;
            }
        }
        //
        internal static int verifyDefaultEmailTemplateGetId(CPBaseClass cp) {
            using (var cs = cp.CSNew()) {
                // 
                // -- try default template
                cs.Open("Newsletter Templates", "name='Newsletter Template Default Email'");
                if (!cs.OK()) {
                    cs.Close();
                    cs.Insert("Newsletter Templates");
                    if (cs.OK()) {
                        cs.SetField("name", "Newsletter Template Default Email");
                    }
                }
                if (cs.OK()) {
                    // 
                    // Use the default template in their Db already
                    if (string.IsNullOrEmpty((cs.GetText("Template") ?? "").Trim())) {
                        cs.SetField("Template", cp.Layout.GetLayout(Constants.guidLayoutDefaultEmailTemplate));
                    }
                    return cs.GetInteger("ID");
                }
                cs.Close();
                return 0;
            }
        }
        // 
        // ===================================================================================================
        // Wrap the content in a common wrapper if authoring is enabled
        // ===================================================================================================
        // 
        public static string GetEditWrapper(CPBaseClass cp, string Caption, string Content) {
            return cp.Content.GetEditWrapper(Content);
        }
        // 
        // ===================================================================================================
        // Wrap the content in a common wrapper if authoring is enabled
        // ===================================================================================================
        // 
        public static string GetAdminHintWrapper(CPBaseClass cp, string Content) {
            return cp.Html.adminHint(Content);
        }
        // 
        internal static DateTime encodeMinDate(DateTime source) {
            var returnDate = source;
            if (returnDate < DateTime.Parse("1/1/1990")) {
                returnDate = DateTime.MinValue;
            }

            return default;
        }
        // 
        // =================================================================================
        // Get a Random Long Value
        // =================================================================================
        // 
        public static int GetRandomInteger() {
            int GetRandomIntegerRet = default;
            int RandomLimit;
            RandomLimit = 32767;
            GetRandomIntegerRet = (int)Math.Round(_random.NextDouble() * RandomLimit);
            return GetRandomIntegerRet;
        }
        // 
        // 
        // 
        internal static bool isBlank(CPBaseClass cp, string source) {
            bool returnBool = false;
            try {
                string test = source;
                // 
                if (test.Length == 0) {
                    returnBool = true;
                } else if (test.Length < 1000) {
                    test = cp.Utils.ConvertHTML2Text(test);
                    test = test.Replace("\n", "");
                    test = test.Replace("\r", "");
                    test = test.Replace("\t", "");
                    test = test.Replace(" ", "");
                    returnBool = test.Length == 0;
                }
            } catch (Exception ex) {
                handleError(cp, ex, "isBlank");
            }
            return returnBool;
        }
    }
}