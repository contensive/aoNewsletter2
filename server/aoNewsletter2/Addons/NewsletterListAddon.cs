using System;
using System.Data;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that lists all newsletters. Clicking a newsletter navigates to its issues.
    /// </summary>
    public class NewsletterListAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterList;
        public const string guidAddon = Constants.guidAddonNewsletterList;
        //
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return "<p>You are not authorized to access this feature.</p>";
                }
                if (!cp.AdminUI.EndpointContainsPortal()) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                processForm(cp);
                return getForm(cp);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static void processForm(CPBaseClass cp) {
            try {
                if (!cp.Doc.IsProperty(Constants.rnButton)) { return; }
                string button = cp.Doc.GetText(Constants.rnButton);
                if (button == Constants.buttonAdd) {
                    using (var cs = cp.CSNew()) {
                        cs.Insert(Constants.ContentNameNewsletters);
                        if (cs.OK()) {
                            int newId = cs.GetInteger("id");
                            cs.SetField("name", $"Newsletter {newId}");
                            cs.SetField("active", "1");
                            cs.Close();
                            string detailLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterDetail) + $"&{Constants.rnNewsletterId}={newId}";
                            cp.Response.Redirect(detailLink);
                        }
                    }
                    return;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static string getForm(CPBaseClass cp) {
            try {
                if (!cp.Response.isOpen) { return ""; }
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
                layoutBuilder.columnCaption = "Issues";
                layoutBuilder.columnCaptionClass = "afwWidth100px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "afwTextAlignCenter";
                //
                layoutBuilder.addColumn();
                layoutBuilder.columnCaption = "Active";
                layoutBuilder.columnCaptionClass = "afwWidth100px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "afwTextAlignCenter";
                //
                // -- sql where clause
                string sqlWhere = "(1=1)";
                if (!string.IsNullOrEmpty(layoutBuilder.sqlSearchTerm)) {
                    sqlWhere += $" and(n.name like {cp.Db.EncodeSQLTextLike(layoutBuilder.sqlSearchTerm)})";
                }
                //
                // -- count
                string sqlCount = $"select count(*) from Newsletters n where {sqlWhere}";
                using (DataTable dt = cp.Db.ExecuteQuery(sqlCount)) {
                    if (dt?.Rows != null && dt.Rows.Count == 1) {
                        layoutBuilder.recordCount = cp.Utils.EncodeInteger(dt.Rows[0][0]);
                    }
                }
                //
                // -- data query
                string sql = $"select n.id, n.name, n.active, (select count(*) from NewsletterIssues i where i.newsletterid=n.id) as issueCount from Newsletters n where {sqlWhere}";
                sql += string.IsNullOrEmpty(layoutBuilder.sqlOrderBy) ? " order by n.name" : $" order by {layoutBuilder.sqlOrderBy}";
                sql += $" OFFSET {(layoutBuilder.paginationPageNumber - 1) * layoutBuilder.paginationPageSize} ROWS FETCH NEXT {layoutBuilder.paginationPageSize} ROWS ONLY";
                //
                string detailLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterDetail) + $"&{Constants.rnNewsletterId}=";
                //
                int rowPtr = 0;
                using (var csList = cp.CSNew()) {
                    if (csList.OpenSQL(sql)) {
                        int rowPtrStart = layoutBuilder.paginationPageSize * (layoutBuilder.paginationPageNumber - 1);
                        do {
                            int newsletterId = csList.GetInteger("id");
                            string newsletterName = csList.GetText("name");
                            if (string.IsNullOrWhiteSpace(newsletterName)) { newsletterName = "(no name)"; }
                            string nameLink = $"<a href=\"{detailLink}{newsletterId}\">{newsletterName}</a>";
                            int issueCount = csList.GetInteger("issueCount");
                            bool isActive = csList.GetBoolean("active");
                            //
                            layoutBuilder.addRow();
                            layoutBuilder.setCell((rowPtrStart + rowPtr + 1).ToString());
                            layoutBuilder.setCell(newsletterId.ToString());
                            layoutBuilder.setCell(nameLink);
                            layoutBuilder.setCell(issueCount.ToString());
                            layoutBuilder.setCell(isActive ? "Yes" : "No");
                            //
                            rowPtr += 1;
                            csList.GoNext();
                        } while (csList.OK());
                        csList.Close();
                    }
                }
                //
                // -- layout settings
                layoutBuilder.title = "Newsletters";
                layoutBuilder.description = "Click a newsletter to see its issues.";
                layoutBuilder.callbackAddonGuid = Constants.guidAddonNewsletterList;
                layoutBuilder.includeBodyColor = true;
                layoutBuilder.includeBodyPadding = true;
                layoutBuilder.includeForm = true;
                layoutBuilder.isOuterContainer = false;
                layoutBuilder.paginationPageSizeDefault = 50;
                //
                // -- buttons
                layoutBuilder.addFormButton(Constants.buttonAdd, Constants.rnButton);
                //
                // -- hiddens
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterList);
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterList);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
