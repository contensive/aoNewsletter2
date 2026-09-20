
using Contensive.Addons.Mcp.Shared;
using Contensive.Addons.Newsletter.Models.Db;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using System;
using System.Collections.Generic;
using System.Linq;
//
namespace Contensive.Addons.Newsletter.Addons {
    //
    /// <summary>
    /// MCP extension addon that provides AI-callable tools for managing newsletter content.
    /// Called by the MCP server with mcpToolName and mcpArguments set as doc properties.
    /// </summary>
    public class NewsletterMcpTools : AddonBaseClass {
        //
        // -- tool name constants
        private const string ToolNewsletterList = "newsletter_list";
        private const string ToolNewsletterGet = "newsletter_get";
        private const string ToolNewsletterUpdate = "newsletter_update";
        private const string ToolIssueList = "newsletter_issue_list";
        private const string ToolIssueGet = "newsletter_issue_get";
        private const string ToolIssueCreate = "newsletter_issue_create";
        private const string ToolIssueUpdate = "newsletter_issue_update";
        private const string ToolIssueDelete = "newsletter_issue_delete";
        private const string ToolStoryList = "newsletter_story_list";
        private const string ToolStoryGet = "newsletter_story_get";
        private const string ToolStoryCreate = "newsletter_story_create";
        private const string ToolStoryUpdate = "newsletter_story_update";
        private const string ToolStoryDelete = "newsletter_story_delete";
        private const string ToolStoryReorder = "newsletter_story_reorder";
        //
        // ====================================================================================================
        //
        public override object Execute(CPBaseClass cp) {
            try {
                string toolName = cp.Doc.GetText("mcpToolName");
                if (string.IsNullOrEmpty(toolName)) {
                    return McpResponseHelper.Error(cp, "This addon is intended for MCP tool calls only.");
                }
                if (!cp.User.IsAdmin) {
                    return McpResponseHelper.Error(cp, "Newsletter MCP tools require admin access.");
                }
                string argsJson = cp.Doc.GetText("mcpArguments");
                var args = string.IsNullOrEmpty(argsJson)
                    ? new Dictionary<string, object>()
                    : cp.JSON.Deserialize<Dictionary<string, object>>(argsJson)
                      ?? new Dictionary<string, object>();
                //
                switch (toolName) {
                    case ToolNewsletterList:
                        return newsletterList(cp, args);
                    case ToolNewsletterGet:
                        return newsletterGet(cp, args);
                    case ToolNewsletterUpdate:
                        return newsletterUpdate(cp, args);
                    case ToolIssueList:
                        return issueList(cp, args);
                    case ToolIssueGet:
                        return issueGet(cp, args);
                    case ToolIssueCreate:
                        return issueCreate(cp, args);
                    case ToolIssueUpdate:
                        return issueUpdate(cp, args);
                    case ToolIssueDelete:
                        return issueDelete(cp, args);
                    case ToolStoryList:
                        return storyList(cp, args);
                    case ToolStoryGet:
                        return storyGet(cp, args);
                    case ToolStoryCreate:
                        return storyCreate(cp, args);
                    case ToolStoryUpdate:
                        return storyUpdate(cp, args);
                    case ToolStoryDelete:
                        return storyDelete(cp, args);
                    case ToolStoryReorder:
                        return storyReorder(cp, args);
                    default:
                        return McpResponseHelper.Error(cp, $"Unknown tool: {toolName}");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Internal error processing newsletter MCP tool.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_list
        // ====================================================================================================
        //
        private string newsletterList(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int pageSize = McpResponseHelper.GetIntArg(args, "pageSize", 50);
                int pageNumber = McpResponseHelper.GetIntArg(args, "pageNumber", 1);
                //
                var newsletters = DbBaseModel.createList<NewsletterModel>(cp, "(active<>0)", "name", pageSize, pageNumber);
                var result = newsletters.Select(n => new {
                    newsletterId = n.id,
                    name = n.name,
                    templateId = n.templateId,
                    emailTemplateId = n.emailTemplateId,
                    blockArchiveSearchForm = n.blockArchiveSearchForm,
                    archiveIssuesToDisplay = n.archiveIssuesToDisplay,
                    searchResultsPerPage = n.searchResultsPerPage
                }).ToList();
                return McpResponseHelper.Success(cp, new { newsletters = result, pageSize, pageNumber }, $"Found {result.Count} newsletter(s)");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error listing newsletters.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_get
        // ====================================================================================================
        //
        private string newsletterGet(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int newsletterId = McpResponseHelper.GetIntArg(args, "newsletterId");
                if (newsletterId == 0) {
                    return McpResponseHelper.Error(cp, "newsletterId is required.");
                }
                var newsletter = DbBaseModel.create<NewsletterModel>(cp, newsletterId);
                if (newsletter == null) {
                    return McpResponseHelper.Error(cp, $"Newsletter #{newsletterId} not found.");
                }
                var result = new {
                    newsletterId = newsletter.id,
                    name = newsletter.name,
                    templateId = newsletter.templateId,
                    emailTemplateId = newsletter.emailTemplateId,
                    blockArchiveSearchForm = newsletter.blockArchiveSearchForm,
                    archiveIssuesToDisplay = newsletter.archiveIssuesToDisplay,
                    searchResultsPerPage = newsletter.searchResultsPerPage,
                    active = newsletter.active
                };
                return McpResponseHelper.Success(cp, result, "OK");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error retrieving newsletter.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_update
        // ====================================================================================================
        //
        private string newsletterUpdate(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int newsletterId = McpResponseHelper.GetIntArg(args, "newsletterId");
                if (newsletterId == 0) {
                    return McpResponseHelper.Error(cp, "newsletterId is required.");
                }
                var newsletter = DbBaseModel.create<NewsletterModel>(cp, newsletterId);
                if (newsletter == null) {
                    return McpResponseHelper.Error(cp, $"Newsletter #{newsletterId} not found.");
                }
                //
                // -- capture undo before changes
                var fieldsBefore = new Dictionary<string, string> {
                    ["name"] = newsletter.name ?? "",
                    ["blockArchiveSearchForm"] = newsletter.blockArchiveSearchForm ? "1" : "0",
                    ["archiveIssuesToDisplay"] = newsletter.archiveIssuesToDisplay.ToString(),
                    ["searchResultsPerPage"] = newsletter.searchResultsPerPage.ToString()
                };
                McpUndoHelper.Capture(cp, Constants.ContentNameNewsletters, newsletter.id, newsletter.ccguid, ToolNewsletterUpdate, fieldsBefore);
                //
                // -- apply updates for any provided fields
                if (args.ContainsKey("name")) { newsletter.name = McpResponseHelper.GetStringArg(args, "name"); }
                if (args.ContainsKey("blockArchiveSearchForm")) { newsletter.blockArchiveSearchForm = McpResponseHelper.GetBoolArg(args, "blockArchiveSearchForm"); }
                if (args.ContainsKey("archiveIssuesToDisplay")) { newsletter.archiveIssuesToDisplay = McpResponseHelper.GetIntArg(args, "archiveIssuesToDisplay"); }
                if (args.ContainsKey("searchResultsPerPage")) { newsletter.searchResultsPerPage = McpResponseHelper.GetIntArg(args, "searchResultsPerPage"); }
                //
                newsletter.save(cp);
                return McpResponseHelper.Success(cp, new { newsletterId = newsletter.id, name = newsletter.name }, "Newsletter updated successfully.");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error updating newsletter.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_issue_list
        // ====================================================================================================
        //
        private string issueList(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int pageSize = McpResponseHelper.GetIntArg(args, "pageSize", 50);
                int pageNumber = McpResponseHelper.GetIntArg(args, "pageNumber", 1);
                int newsletterId = McpResponseHelper.GetIntArg(args, "newsletterId");
                //
                var criteria = new List<string> { "(active<>0)" };
                if (newsletterId > 0) {
                    criteria.Add($"(newsletterid={newsletterId})");
                }
                string where = string.Join(" AND ", criteria);
                //
                var issues = new List<object>();
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (cs.Open(Constants.ContentNameNewsletterIssues, where, "publishdate desc", true, "id,name,newsletterid,publishdate,sponsor,tagline")) {
                        int skip = (pageNumber - 1) * pageSize;
                        int skipped = 0;
                        int taken = 0;
                        do {
                            if (skipped < skip) {
                                skipped++;
                                cs.GoNext();
                                continue;
                            }
                            if (taken >= pageSize) { break; }
                            DateTime pubDate = cs.GetDate("publishdate");
                            issues.Add(new {
                                issueId = cs.GetInteger("id"),
                                name = cs.GetText("name"),
                                newsletterId = cs.GetInteger("newsletterid"),
                                publishDate = pubDate == DateTime.MinValue ? "" : pubDate.ToString("yyyy-MM-dd"),
                                sponsor = cs.GetText("sponsor"),
                                tagline = cs.GetText("tagline")
                            });
                            taken++;
                            cs.GoNext();
                        } while (cs.OK());
                    }
                }
                return McpResponseHelper.Success(cp, new { issues, pageSize, pageNumber }, $"Found {issues.Count} issue(s)");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error listing newsletter issues.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_issue_get
        // ====================================================================================================
        //
        private string issueGet(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int issueId = McpResponseHelper.GetIntArg(args, "issueId");
                if (issueId == 0) {
                    return McpResponseHelper.Error(cp, "issueId is required.");
                }
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}")) {
                        return McpResponseHelper.Error(cp, $"Issue #{issueId} not found.");
                    }
                    DateTime pubDate = cs.GetDate("publishdate");
                    var result = new {
                        issueId = cs.GetInteger("id"),
                        name = cs.GetText("name"),
                        newsletterId = cs.GetInteger("newsletterid"),
                        publishDate = pubDate == DateTime.MinValue ? "" : pubDate.ToString("yyyy-MM-dd"),
                        cover = cs.GetText("cover"),
                        overview = cs.GetText("overview"),
                        sponsor = cs.GetText("sponsor"),
                        tagline = cs.GetText("tagline"),
                        active = cs.GetBoolean("active"),
                        sortOrder = cs.GetText("sortorder")
                    };
                    return McpResponseHelper.Success(cp, result, "OK");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error retrieving newsletter issue.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_issue_create
        // ====================================================================================================
        //
        private string issueCreate(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int newsletterId = McpResponseHelper.GetIntArg(args, "newsletterId");
                string name = McpResponseHelper.GetStringArg(args, "name");
                if (newsletterId == 0) {
                    return McpResponseHelper.Error(cp, "newsletterId is required.");
                }
                if (string.IsNullOrEmpty(name)) {
                    return McpResponseHelper.Error(cp, "name is required.");
                }
                //
                // -- verify the newsletter exists
                var newsletter = DbBaseModel.create<NewsletterModel>(cp, newsletterId);
                if (newsletter == null) {
                    return McpResponseHelper.Error(cp, $"Newsletter #{newsletterId} not found.");
                }
                //
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Insert(Constants.ContentNameNewsletterIssues)) {
                        return McpResponseHelper.Error(cp, "Error creating issue record.");
                    }
                    int newId = cs.GetInteger("id");
                    string ccguid = cs.GetText("ccguid");
                    cs.SetField("name", name);
                    cs.SetField("newsletterid", newsletterId.ToString());
                    cs.SetField("active", "1");
                    //
                    // -- optional fields
                    string dateStr = McpResponseHelper.GetStringArg(args, "publishDate");
                    if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out DateTime parsedDate)) {
                        cs.SetField("publishdate", parsedDate.ToString());
                    }
                    if (args.ContainsKey("sponsor")) { cs.SetField("sponsor", McpResponseHelper.GetStringArg(args, "sponsor")); }
                    if (args.ContainsKey("cover")) { cs.SetField("cover", McpResponseHelper.GetStringArg(args, "cover")); }
                    if (args.ContainsKey("overview")) { cs.SetField("overview", McpResponseHelper.GetStringArg(args, "overview")); }
                    if (args.ContainsKey("tagline")) { cs.SetField("tagline", McpResponseHelper.GetStringArg(args, "tagline")); }
                    cs.Save();
                    //
                    // -- capture undo (recordDeleted=true means undo will delete this newly created record)
                    McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterIssues, newId, ccguid, ToolIssueCreate, new Dictionary<string, string>(), recordDeleted: true);
                    //
                    return McpResponseHelper.Success(cp, new { issueId = newId, name, newsletterId }, "Issue created successfully.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error creating newsletter issue.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_issue_update
        // ====================================================================================================
        //
        private string issueUpdate(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int issueId = McpResponseHelper.GetIntArg(args, "issueId");
                if (issueId == 0) {
                    return McpResponseHelper.Error(cp, "issueId is required.");
                }
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}")) {
                        return McpResponseHelper.Error(cp, $"Issue #{issueId} not found.");
                    }
                    //
                    // -- capture undo before changes
                    DateTime pubDateBefore = cs.GetDate("publishdate");
                    var fieldsBefore = new Dictionary<string, string> {
                        ["name"] = cs.GetText("name"),
                        ["cover"] = cs.GetText("cover"),
                        ["overview"] = cs.GetText("overview"),
                        ["publishdate"] = pubDateBefore == DateTime.MinValue ? "" : pubDateBefore.ToString("o"),
                        ["sponsor"] = cs.GetText("sponsor"),
                        ["tagline"] = cs.GetText("tagline"),
                        ["sortorder"] = cs.GetText("sortorder")
                    };
                    string ccguid = cs.GetText("ccguid");
                    McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterIssues, issueId, ccguid, ToolIssueUpdate, fieldsBefore);
                    //
                    // -- apply updates for any provided fields
                    if (args.ContainsKey("name")) { cs.SetField("name", McpResponseHelper.GetStringArg(args, "name")); }
                    if (args.ContainsKey("cover")) { cs.SetField("cover", McpResponseHelper.GetStringArg(args, "cover")); }
                    if (args.ContainsKey("overview")) { cs.SetField("overview", McpResponseHelper.GetStringArg(args, "overview")); }
                    if (args.ContainsKey("sponsor")) { cs.SetField("sponsor", McpResponseHelper.GetStringArg(args, "sponsor")); }
                    if (args.ContainsKey("tagline")) { cs.SetField("tagline", McpResponseHelper.GetStringArg(args, "tagline")); }
                    if (args.ContainsKey("sortOrder")) { cs.SetField("sortorder", McpResponseHelper.GetStringArg(args, "sortOrder")); }
                    if (args.ContainsKey("publishDate")) {
                        string dateStr = McpResponseHelper.GetStringArg(args, "publishDate");
                        if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out DateTime parsedDate)) {
                            cs.SetField("publishdate", parsedDate.ToString());
                        }
                    }
                    cs.Save();
                    return McpResponseHelper.Success(cp, new { issueId, name = cs.GetText("name") }, "Issue updated successfully.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error updating newsletter issue.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_issue_delete
        // ====================================================================================================
        //
        private string issueDelete(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int issueId = McpResponseHelper.GetIntArg(args, "issueId");
                if (issueId == 0) {
                    return McpResponseHelper.Error(cp, "issueId is required.");
                }
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}")) {
                        return McpResponseHelper.Error(cp, $"Issue #{issueId} not found.");
                    }
                    string name = cs.GetText("name");
                    string ccguid = cs.GetText("ccguid");
                    var fieldsBefore = new Dictionary<string, string> {
                        ["active"] = cs.GetBoolean("active").ToString()
                    };
                    McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterIssues, issueId, ccguid, ToolIssueDelete, fieldsBefore);
                    //
                    cs.SetField("active", "0");
                    cs.Save();
                    return McpResponseHelper.Success(cp, new { issueId, name }, "Issue deleted successfully.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error deleting newsletter issue.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_story_list
        // ====================================================================================================
        //
        private string storyList(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int pageSize = McpResponseHelper.GetIntArg(args, "pageSize", 50);
                int pageNumber = McpResponseHelper.GetIntArg(args, "pageNumber", 1);
                int issueId = McpResponseHelper.GetIntArg(args, "issueId");
                //
                // -- note: the "newsletterid" column in NewsletterIssuePages references the issue id, not the newsletter id
                var criteria = new List<string> { "(active<>0)" };
                if (issueId > 0) {
                    criteria.Add($"(newsletterid={issueId})");
                }
                string where = string.Join(" AND ", criteria);
                //
                var stories = new List<object>();
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (cs.Open(Constants.ContentNameNewsletterStories, where, "sortorder,name", true, "id,name,newsletterid,sortorder,categoryid,active")) {
                        int skip = (pageNumber - 1) * pageSize;
                        int skipped = 0;
                        int taken = 0;
                        do {
                            if (skipped < skip) {
                                skipped++;
                                cs.GoNext();
                                continue;
                            }
                            if (taken >= pageSize) { break; }
                            stories.Add(new {
                                storyId = cs.GetInteger("id"),
                                name = cs.GetText("name"),
                                issueId = cs.GetInteger("newsletterid"),
                                sortOrder = cs.GetText("sortorder"),
                                categoryId = cs.GetInteger("categoryid"),
                                active = cs.GetBoolean("active")
                            });
                            taken++;
                            cs.GoNext();
                        } while (cs.OK());
                    }
                }
                return McpResponseHelper.Success(cp, new { stories, pageSize, pageNumber }, $"Found {stories.Count} story/stories");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error listing newsletter stories.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_story_get
        // ====================================================================================================
        //
        private string storyGet(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int storyId = McpResponseHelper.GetIntArg(args, "storyId");
                if (storyId == 0) {
                    return McpResponseHelper.Error(cp, "storyId is required.");
                }
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId}")) {
                        return McpResponseHelper.Error(cp, $"Story #{storyId} not found.");
                    }
                    var result = new {
                        storyId = cs.GetInteger("id"),
                        name = cs.GetText("name"),
                        overview = cs.GetText("overview"),
                        body = cs.GetText("body"),
                        issueId = cs.GetInteger("newsletterid"),
                        categoryId = cs.GetInteger("categoryid"),
                        sortOrder = cs.GetText("sortorder"),
                        active = cs.GetBoolean("active")
                    };
                    return McpResponseHelper.Success(cp, result, "OK");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error retrieving newsletter story.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_story_create
        // ====================================================================================================
        //
        private string storyCreate(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int issueId = McpResponseHelper.GetIntArg(args, "issueId");
                string name = McpResponseHelper.GetStringArg(args, "name");
                if (issueId == 0) {
                    return McpResponseHelper.Error(cp, "issueId is required.");
                }
                if (string.IsNullOrEmpty(name)) {
                    return McpResponseHelper.Error(cp, "name is required.");
                }
                //
                // -- verify the issue exists
                using (CPCSBaseClass csCheck = cp.CSNew()) {
                    if (!csCheck.Open(Constants.ContentNameNewsletterIssues, $"id={issueId}", "", false, "id")) {
                        return McpResponseHelper.Error(cp, $"Issue #{issueId} not found.");
                    }
                }
                //
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Insert(Constants.ContentNameNewsletterStories)) {
                        return McpResponseHelper.Error(cp, "Error creating story record.");
                    }
                    int newId = cs.GetInteger("id");
                    string ccguid = cs.GetText("ccguid");
                    cs.SetField("name", name);
                    // -- note: the "newsletterid" column in NewsletterIssuePages references the issue id
                    cs.SetField("newsletterid", issueId.ToString());
                    cs.SetField("active", "1");
                    //
                    // -- optional fields
                    if (args.ContainsKey("overview")) { cs.SetField("overview", McpResponseHelper.GetStringArg(args, "overview")); }
                    if (args.ContainsKey("body")) { cs.SetField("body", McpResponseHelper.GetStringArg(args, "body")); }
                    if (args.ContainsKey("categoryId")) { cs.SetField("categoryid", McpResponseHelper.GetIntArg(args, "categoryId").ToString()); }
                    if (args.ContainsKey("sortOrder")) { cs.SetField("sortorder", McpResponseHelper.GetStringArg(args, "sortOrder")); }
                    cs.Save();
                    //
                    // -- capture undo (recordDeleted=true means undo will delete this newly created record)
                    McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterStories, newId, ccguid, ToolStoryCreate, new Dictionary<string, string>(), recordDeleted: true);
                    //
                    return McpResponseHelper.Success(cp, new { storyId = newId, name, issueId }, "Story created successfully.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error creating newsletter story.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_story_update
        // ====================================================================================================
        //
        private string storyUpdate(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int storyId = McpResponseHelper.GetIntArg(args, "storyId");
                if (storyId == 0) {
                    return McpResponseHelper.Error(cp, "storyId is required.");
                }
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId}")) {
                        return McpResponseHelper.Error(cp, $"Story #{storyId} not found.");
                    }
                    //
                    // -- capture undo before changes
                    var fieldsBefore = new Dictionary<string, string> {
                        ["name"] = cs.GetText("name"),
                        ["overview"] = cs.GetText("overview"),
                        ["body"] = cs.GetText("body"),
                        ["categoryid"] = cs.GetInteger("categoryid").ToString(),
                        ["sortorder"] = cs.GetText("sortorder")
                    };
                    string ccguid = cs.GetText("ccguid");
                    McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterStories, storyId, ccguid, ToolStoryUpdate, fieldsBefore);
                    //
                    // -- apply updates for any provided fields
                    if (args.ContainsKey("name")) { cs.SetField("name", McpResponseHelper.GetStringArg(args, "name")); }
                    if (args.ContainsKey("overview")) { cs.SetField("overview", McpResponseHelper.GetStringArg(args, "overview")); }
                    if (args.ContainsKey("body")) { cs.SetField("body", McpResponseHelper.GetStringArg(args, "body")); }
                    if (args.ContainsKey("categoryId")) { cs.SetField("categoryid", McpResponseHelper.GetIntArg(args, "categoryId").ToString()); }
                    if (args.ContainsKey("sortOrder")) { cs.SetField("sortorder", McpResponseHelper.GetStringArg(args, "sortOrder")); }
                    cs.Save();
                    return McpResponseHelper.Success(cp, new { storyId, name = cs.GetText("name") }, "Story updated successfully.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error updating newsletter story.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_story_delete
        // ====================================================================================================
        //
        private string storyDelete(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int storyId = McpResponseHelper.GetIntArg(args, "storyId");
                if (storyId == 0) {
                    return McpResponseHelper.Error(cp, "storyId is required.");
                }
                using (CPCSBaseClass cs = cp.CSNew()) {
                    if (!cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId}")) {
                        return McpResponseHelper.Error(cp, $"Story #{storyId} not found.");
                    }
                    string name = cs.GetText("name");
                    string ccguid = cs.GetText("ccguid");
                    var fieldsBefore = new Dictionary<string, string> {
                        ["active"] = cs.GetBoolean("active").ToString()
                    };
                    McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterStories, storyId, ccguid, ToolStoryDelete, fieldsBefore);
                    //
                    cs.SetField("active", "0");
                    cs.Save();
                    return McpResponseHelper.Success(cp, new { storyId, name }, "Story deleted successfully.");
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error deleting newsletter story.");
            }
        }
        //
        // ====================================================================================================
        // -- newsletter_story_reorder
        // ====================================================================================================
        //
        private string storyReorder(CPBaseClass cp, Dictionary<string, object> args) {
            try {
                int issueId = McpResponseHelper.GetIntArg(args, "issueId");
                string storyIdOrder = McpResponseHelper.GetStringArg(args, "storyIdOrder");
                if (issueId == 0) {
                    return McpResponseHelper.Error(cp, "issueId is required.");
                }
                if (string.IsNullOrEmpty(storyIdOrder)) {
                    return McpResponseHelper.Error(cp, "storyIdOrder is required.");
                }
                //
                string[] idStrings = storyIdOrder.Split(',');
                int sortIndex = 0;
                foreach (string idStr in idStrings) {
                    if (!int.TryParse(idStr.Trim(), out int storyId) || storyId == 0) { continue; }
                    // -- note: the "newsletterid" column in NewsletterIssuePages references the issue id
                    using (CPCSBaseClass cs = cp.CSNew()) {
                        if (cs.Open(Constants.ContentNameNewsletterStories, $"id={storyId} and newsletterid={issueId}")) {
                            string ccguid = cs.GetText("ccguid");
                            var fieldsBefore = new Dictionary<string, string> {
                                ["sortorder"] = cs.GetText("sortorder")
                            };
                            McpUndoHelper.Capture(cp, Constants.ContentNameNewsletterStories, storyId, ccguid, ToolStoryReorder, fieldsBefore);
                            //
                            // -- use zero-padded numbers so alpha sort works: "0010", "0020", etc.
                            string newSortOrder = ((sortIndex + 1) * 10).ToString("D4");
                            cs.SetField("sortorder", newSortOrder);
                            cs.Save();
                        }
                    }
                    sortIndex++;
                }
                return McpResponseHelper.Success(cp, new { issueId, storiesReordered = sortIndex }, $"Reordered {sortIndex} story/stories.");
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return McpResponseHelper.Error(cp, "Error reordering newsletter stories.");
            }
        }
    }
}
