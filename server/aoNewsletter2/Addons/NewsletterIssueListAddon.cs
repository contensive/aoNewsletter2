using System;
using System.Data;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that lists the issues for a specific newsletter.
    /// </summary>
    public class NewsletterIssueListAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterIssueList;
        public const string guidAddon = Constants.guidAddonNewsletterIssueList;
        //
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return "<p>You are not authorized to access this feature.</p>";
                }
                if (!cp.AdminUI.EndpointContainsPortal()) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                int newsletterId = cp.Doc.GetInteger(Constants.rnNewsletterId);
                if (newsletterId == 0) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                processForm(cp, newsletterId);
                return getForm(cp, newsletterId);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static void processForm(CPBaseClass cp, int newsletterId) {
            try {
                if (!cp.Doc.IsProperty(Constants.rnButton)) { return; }
                string button = cp.Doc.GetText(Constants.rnButton);
                if (button == Constants.buttonAdd) {
                    using (var cs = cp.CSNew()) {
                        cs.Insert(Constants.ContentNameNewsletterIssues);
                        if (cs.OK()) {
                            int newId = cs.GetInteger("id");
                            cs.SetField("name", $"Issue {newId}");
                            cs.SetField("newsletterid", newsletterId.ToString());
                            cs.SetField("active", "1");
                            cs.Close();
                        }
                    }
                    return;
                }
                if (button == Constants.buttonBack) {
                    cp.Response.Redirect(cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList));
                    return;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static string getForm(CPBaseClass cp, int newsletterId) {
            try {
                if (!cp.Response.isOpen) { return ""; }
                //
                // -- get newsletter name for title
                string newsletterName = "";
                using (var cs = cp.CSNew()) {
                    cs.Open(Constants.ContentNameNewsletters, $"id={newsletterId}", "", false, "name");
                    if (cs.OK()) {
                        newsletterName = cs.GetText("name");
                    }
                    cs.Close();
                }
                //
                var layoutBuilder = cp.AdminUI.CreateLayoutBuilderList();
                //
                // -- columns
                layoutBuilder.columnCaption = "Row";
                layoutBuilder.columnCaptionClass = "afwWidth20px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "";
                //
                layoutBuilder.addColumn();
                layoutBuilder.columnCaption = "ID";
                layoutBuilder.columnCaptionClass = "afwWidth20px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "";
                //
                layoutBuilder.addColumn();
                layoutBuilder.columnCaption = "Name";
                layoutBuilder.columnCaptionClass = "afwTextAlignLeft";
                layoutBuilder.columnCellClass = "afwTextAlignLeft";
                layoutBuilder.columnSortable = false;
                //
                layoutBuilder.addColumn();
                layoutBuilder.columnCaption = "Publish Date";
                layoutBuilder.columnCaptionClass = "afwWidth100px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "afwTextAlignCenter";
                //
                layoutBuilder.addColumn();
                layoutBuilder.columnCaption = "Stories";
                layoutBuilder.columnCaptionClass = "afwWidth100px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "afwTextAlignCenter";
                //
                // -- sql where clause
                string sqlWhere = $"(i.newsletterid={cp.Db.EncodeSQLNumber(newsletterId)})";
                if (!string.IsNullOrEmpty(layoutBuilder.sqlSearchTerm)) {
                    sqlWhere += $" and(i.name like {cp.Db.EncodeSQLTextLike(layoutBuilder.sqlSearchTerm)})";
                }
                //
                // -- count
                string sqlCount = $"select count(*) from NewsletterIssues i where {sqlWhere}";
                using (DataTable dt = cp.Db.ExecuteQuery(sqlCount)) {
                    if (dt?.Rows != null && dt.Rows.Count == 1) {
                        layoutBuilder.recordCount = cp.Utils.EncodeInteger(dt.Rows[0][0]);
                    }
                }
                //
                // -- data query
                string sql = $"select i.id, i.name, i.publishdate, (select count(*) from NewsletterIssuePages s where s.newsletterid=i.id) as storyCount from NewsletterIssues i where {sqlWhere}";
                sql += string.IsNullOrEmpty(layoutBuilder.sqlOrderBy) ? " order by i.publishdate desc" : $" order by {layoutBuilder.sqlOrderBy}";
                sql += $" OFFSET {(layoutBuilder.paginationPageNumber - 1) * layoutBuilder.paginationPageSize} ROWS FETCH NEXT {layoutBuilder.paginationPageSize} ROWS ONLY";
                //
                int rowPtr = 0;
                int issueCid = cp.Content.GetID(Constants.ContentNameNewsletterIssues);
                using (var csList = cp.CSNew()) {
                    if (csList.OpenSQL(sql)) {
                        int rowPtrStart = layoutBuilder.paginationPageSize * (layoutBuilder.paginationPageNumber - 1);
                        do {
                            int issueId = csList.GetInteger("id");
                            string issueName = csList.GetText("name");
                            if (string.IsNullOrWhiteSpace(issueName)) { issueName = "(no name)"; }
                            string editLink = $"?af=4&aa=2&ad=1&cid={issueCid}&id={issueId}";
                            issueName = $"<a href=\"{editLink}\">{issueName}</a>";
                            DateTime publishDate = csList.GetDate("publishdate");
                            string publishDateStr = publishDate.Equals(DateTime.MinValue) ? "" : publishDate.ToShortDateString();
                            int storyCount = csList.GetInteger("storyCount");
                            //
                            layoutBuilder.addRow();
                            layoutBuilder.setCell((rowPtrStart + rowPtr + 1).ToString());
                            layoutBuilder.setCell(issueId.ToString());
                            layoutBuilder.setCell(issueName);
                            layoutBuilder.setCell(publishDateStr);
                            layoutBuilder.setCell(storyCount.ToString());
                            //
                            rowPtr += 1;
                            csList.GoNext();
                        } while (csList.OK());
                        csList.Close();
                    }
                }
                //
                // -- layout settings
                layoutBuilder.title = string.IsNullOrWhiteSpace(newsletterName) ? "Newsletter Issues" : $"Issues for: {newsletterName}";
                layoutBuilder.description = "Click an issue to edit it. Use Add to create a new issue for this newsletter.";
                layoutBuilder.callbackAddonGuid = Constants.guidAddonNewsletterIssueList;
                layoutBuilder.includeBodyColor = true;
                layoutBuilder.includeBodyPadding = true;
                layoutBuilder.includeForm = true;
                layoutBuilder.isOuterContainer = false;
                layoutBuilder.paginationPageSizeDefault = 50;
                //
                // -- buttons
                layoutBuilder.addFormButton(Constants.buttonAdd, Constants.rnButton);
                layoutBuilder.addFormButton(Constants.buttonBack, Constants.rnButton);
                //
                // -- hiddens
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterIssueList);
                layoutBuilder.addFormHidden(Constants.rnNewsletterId, newsletterId);
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnNewsletterId, newsletterId);
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterIssueList);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
