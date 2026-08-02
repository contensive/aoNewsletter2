using System;
using System.Data;
using Contensive.BaseClasses;

namespace Contensive.Addons.Newsletter.Addons {
    /// <summary>
    /// Portal addon that lists all stories for a newsletter issue.
    /// Clicking a story navigates to the Story Detail view.
    /// </summary>
    public class NewsletterIssueStoriesAddon : AddonBaseClass {
        //
        public const string guidPortalFeature = Constants.guidPortalFeatureNewsletterIssueStories;
        public const string guidAddon = Constants.guidAddonNewsletterIssueStories;
        //
        public override object Execute(CPBaseClass cp) {
            try {
                if (!cp.User.IsAdmin) {
                    return "<p>You are not authorized to access this feature.</p>";
                }
                if (!cp.AdminUI.EndpointContainsPortal()) {
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                int issueId = cp.Doc.GetInteger(Constants.rnIssueId);
                int newsletterId = cp.Doc.GetInteger(Constants.rnNewsletterId);
                if (issueId == 0) {
                    if (newsletterId != 0) {
                        return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueList, $"{Constants.rnNewsletterId}={newsletterId}");
                    }
                    return cp.AdminUI.RedirectToPortalFeature(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterList, "");
                }
                if (newsletterId == 0) {
                    using (var cs = cp.CSNew()) {
                        cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}", "", false, "newsletterid");
                        if (cs.OK()) {
                            newsletterId = cs.GetInteger("newsletterid");
                        }
                        cs.Close();
                    }
                }
                processForm(cp, issueId, newsletterId);
                return getForm(cp, issueId, newsletterId);
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static void processForm(CPBaseClass cp, int issueId, int newsletterId) {
            try {
                if (!cp.Doc.IsProperty(Constants.rnButton)) { return; }
                string button = cp.Doc.GetText(Constants.rnButton);
                if (button == Constants.buttonAdd) {
                    using (var cs = cp.CSNew()) {
                        cs.Insert(Constants.ContentNameNewsletterStories);
                        if (cs.OK()) {
                            int newStoryId = cs.GetInteger("id");
                            cs.SetField("name", $"Story {newStoryId}");
                            cs.SetField("newsletterid", issueId.ToString());
                            cs.SetField("active", "1");
                            cs.Close();
                            string detailLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterStoryDetail)
                                + $"&{Constants.rnStoryId}={newStoryId}&{Constants.rnIssueId}={issueId}&{Constants.rnNewsletterId}={newsletterId}";
                            cp.Response.Redirect(detailLink);
                        }
                    }
                    return;
                }
                if (button == Constants.buttonBack) {
                    string issueDetailLink = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterIssueDetail)
                        + $"&{Constants.rnIssueId}={issueId}&{Constants.rnNewsletterId}={newsletterId}";
                    cp.Response.Redirect(issueDetailLink);
                    return;
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        //
        internal static string getForm(CPBaseClass cp, int issueId, int newsletterId) {
            try {
                if (!cp.Response.isOpen) { return ""; }
                //
                // -- get newsletter name for sub-nav title
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
                layoutBuilder.portalSubNavTitle = $"{newsletterName}, #{newsletterId}";
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
                layoutBuilder.columnCaption = "Sort Order";
                layoutBuilder.columnCaptionClass = "afwWidth100px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "afwTextAlignCenter";
                //
                layoutBuilder.addColumn();
                layoutBuilder.columnCaption = "Active";
                layoutBuilder.columnCaptionClass = "afwWidth100px afwTextAlignCenter";
                layoutBuilder.columnCellClass = "afwTextAlignCenter";
                //
                // -- sql where clause
                string sqlWhere = $"(s.newsletterid={cp.Db.EncodeSQLNumber(issueId)})";
                if (!string.IsNullOrEmpty(layoutBuilder.sqlSearchTerm)) {
                    sqlWhere += $" and(s.name like {cp.Db.EncodeSQLTextLike(layoutBuilder.sqlSearchTerm)})";
                }
                //
                // -- count
                string sqlCount = $"select count(*) from NewsletterIssuePages s where {sqlWhere}";
                using (DataTable dt = cp.Db.ExecuteQuery(sqlCount)) {
                    if (dt?.Rows != null && dt.Rows.Count == 1) {
                        layoutBuilder.recordCount = cp.Utils.EncodeInteger(dt.Rows[0][0]);
                    }
                }
                //
                // -- data query
                string sql = $"select s.id, s.name, s.sortorder, s.active from NewsletterIssuePages s where {sqlWhere}";
                sql += string.IsNullOrEmpty(layoutBuilder.sqlOrderBy) ? " order by s.sortorder, s.name" : $" order by {layoutBuilder.sqlOrderBy}";
                sql += $" OFFSET {(layoutBuilder.paginationPageNumber - 1) * layoutBuilder.paginationPageSize} ROWS FETCH NEXT {layoutBuilder.paginationPageSize} ROWS ONLY";
                //
                string detailLinkBase = cp.AdminUI.GetPortalFeatureLink(Constants.guidPortalShare, Constants.guidPortalFeatureNewsletterStoryDetail);
                int rowPtr = 0;
                using (var csList = cp.CSNew()) {
                    if (csList.OpenSQL(sql)) {
                        int rowPtrStart = layoutBuilder.paginationPageSize * (layoutBuilder.paginationPageNumber - 1);
                        do {
                            int storyId = csList.GetInteger("id");
                            string storyName = csList.GetText("name");
                            if (string.IsNullOrWhiteSpace(storyName)) { storyName = "(no name)"; }
                            string detailLink = $"{detailLinkBase}&{Constants.rnStoryId}={storyId}&{Constants.rnIssueId}={issueId}&{Constants.rnNewsletterId}={newsletterId}";
                            string nameLink = $"<a href=\"{detailLink}\">{storyName}</a>";
                            string sortOrder = csList.GetText("sortorder");
                            bool isActive = csList.GetBoolean("active");
                            //
                            layoutBuilder.addRow();
                            layoutBuilder.setCell((rowPtrStart + rowPtr + 1).ToString());
                            layoutBuilder.setCell(storyId.ToString());
                            layoutBuilder.setCell(nameLink);
                            layoutBuilder.setCell(sortOrder);
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
                layoutBuilder.title = "Stories";
                layoutBuilder.description = "Click a story to edit it.";
                layoutBuilder.callbackAddonGuid = Constants.guidAddonNewsletterIssueStories;
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
                layoutBuilder.addFormHidden(Constants.rnSrcFormId, Constants.formIdNewsletterIssueStories);
                layoutBuilder.addFormHidden(Constants.rnIssueId, issueId);
                layoutBuilder.addFormHidden(Constants.rnNewsletterId, newsletterId);
                //
                // -- refresh query string
                cp.Doc.AddRefreshQueryString(Constants.rnIssueId, issueId);
                cp.Doc.AddRefreshQueryString(Constants.rnNewsletterId, newsletterId);
                cp.Doc.AddRefreshQueryString(Constants.rnDstFeatureGuid, Constants.guidPortalFeatureNewsletterIssueStories);
                //
                return layoutBuilder.getHtml();
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
    }
}
