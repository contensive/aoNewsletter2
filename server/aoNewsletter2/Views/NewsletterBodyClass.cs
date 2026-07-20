
using System;
using Contensive.Addons.Newsletter.Controllers;
using Contensive.Addons.Newsletter.Models.Db;
using Contensive.BaseClasses;
using System.Globalization;

namespace Contensive.Addons.Newsletter.Views {
    public class NewsletterBodyClass {
        // 
        internal string GetArchiveItemList(CPBaseClass cp, NewsletterController cn, string ButtonValue, int currentIssueId, string refreshQueryString, string newsArchiveListItemLayout, int NewsletterID) {
            string GetArchiveItemListRet = default;
            // 
            var layout = new BlockClass();
            int recordTop;
            int RecordsPerPage;
            int archiveIssuesToDisplay;
            var cs = cp.CSNew();
            int monthSelected;
            int yearSelected;
            string SearchKeywords;
            // 
            string link = "";
            string Stream = "";
            string ThisSQL;
            string ThisSQL2 = "";
            string MonthString;
            string YearString;
            int MonthCounter;
            int YearCounter;
            string storyName;
            int NumberofPages;
            string sql2;
            int FileCount;
            string storyOverview;
            var RowCount = default(int);
            int YearsWanted;
            bool BlockSearchForm;
            string qs;
            int PageNumber;
            DateTime issueDate;
            string issueDateFormatted = "";
            string storyBody;
            //
            BlockSearchForm = true;
            var newsletter = Contensive.Models.Db.DbBaseModel.create<NewsletterModel>(cp, NewsletterID);
            archiveIssuesToDisplay = newsletter.archiveIssuesToDisplay;
            monthSelected = cp.Doc.GetInteger(Constants.RequestNameMonthSelectd);
            yearSelected = cp.Doc.GetInteger(Constants.RequestNameYearSelected);
            SearchKeywords = cp.Doc.GetText(Constants.RequestNameSearchKeywords);
            RecordsPerPage = 10;
            recordTop = cp.Doc.GetInteger(Constants.RequestNameRecordTop);
            // 
            PageNumber = cp.Doc.GetInteger(Constants.RequestNamePageNumber);
            if (PageNumber == 0) {
                PageNumber = 1;
            }
            // 
            YearsWanted = cp.Utils.EncodeInteger(cp.Site.GetText("Newsletter years wanted", "1"));
            if (YearsWanted < 1) {
                YearsWanted = 1;
            }
            // 
            if (archiveIssuesToDisplay == 0) {
                archiveIssuesToDisplay = 6;
            }
            //
            // Get Total Archive count
            //
            sql2 = " select count(story.id) as count";
            sql2 = sql2 + " from newsletterissues nl, newsletterissuepages story";
            sql2 = sql2 + " Where (NL.ID = story.newsletterid)";
            sql2 = sql2 + " AND nl.active=1 and (NL.NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ")";
            if (monthSelected != 0) {
                ThisSQL2 = ThisSQL2 + " and month(nl.publishdate) = " + monthSelected;
            }
            if (yearSelected != 0) {
                ThisSQL2 = ThisSQL2 + " and year(nl.publishdate) = " + yearSelected;
            }
            if (!string.IsNullOrEmpty(SearchKeywords)) {
                sql2 = sql2 + " and ((story.Body like '%" + SearchKeywords + "%' )or (story.name  like '%" + SearchKeywords + "%') or (story.Overview  like '%" + SearchKeywords + "%'))";
            }
            if (cs.OpenSQL(sql2)) {
                FileCount = cs.GetInteger("count");
                NumberofPages = (int)Math.Round(FileCount / (double)RecordsPerPage);
                if (NumberofPages != (int)Math.Floor((double)NumberofPages)) {
                    NumberofPages = NumberofPages + 1;
                    NumberofPages = (int)Math.Floor((double)NumberofPages);
                }
                if (NumberofPages == 0) {
                    NumberofPages = 1;
                }
            }
            cs.Close();
            //
            // Colors = "#ffffff"
            //
            //
            if ((ButtonValue ?? "") != Constants.FormButtonViewNewsLetter & (ButtonValue ?? "") != Constants.FormButtonViewArchives) {
                //
                // List a page of archive issues
                //
                if (monthSelected == 0 & yearSelected == 0) {
                    // stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                    //
                    // ThisSQL = " SELECT  TOP 6 * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & cp.db.encodesqlNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    ThisSQL = " SELECT  TOP " + archiveIssuesToDisplay + " * " + " From NewsletterIssues " + " WHERE active=1 and (PublishDate < { fn NOW() }) AND (ID <> " + currentIssueId + ") AND (NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ") " + " ORDER BY PublishDate DESC";


                    //
                    cs.OpenSQL(ThisSQL);
                    if (cs.OK()) {
                        while (cs.OK()) {
                            layout.load(newsArchiveListItemLayout);
                            issueDate = cs.GetDate("PublishDate");
                            if (issueDate != DateTime.MinValue) {
                                issueDateFormatted = $"{CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(issueDate.Month)} {issueDate.Day}, {issueDate.Year}";
                            }
                            link = refreshQueryString;
                            link = cp.Utils.ModifyQueryString(link, Constants.RequestNameIssueID, cs.GetInteger("ID").ToString());
                            layout.setClassInner("newsArchiveListCaption", cs.GetText("Name"));
                            layout.setClassInner("newsArchiveListOverview", cp.Utils.EncodeContentForWeb(cs.GetText("Overview")));
                            // Stream &= layout.GetHtml().Replace("?", "?" & link)
                            Stream += layout.getHtml().Replace("href=\"?\"", $"href=\"?{link}\"");
                            cs.GoNext();
                        }
                    } else {
                        BlockSearchForm = true;
                        layout.load(newsArchiveListItemLayout);
                        layout.setClassInner("newsArchiveListCaption", "<span class=\"ccError\">" + cp.Site.GetText(Constants.SitePropertyNoNewsletterArchives, "There are currently no archived issues.") + "</span>");
                        layout.setClassInner("newsArchiveListOverview", "");
                        Stream += layout.getHtml();
                    }
                    cs.Close();
                }
            }
            if ((ButtonValue ?? "") == Constants.FormButtonViewArchives) {
                // 
                // List search results of archive issues
                // 
                cp.Utils.AppendLog(cp.Doc.GetInteger("newsletter").ToString());

                // stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                ThisSQL2 = " select NL.id, nl.name, nl.publishdate, story.AllowReadMore, story.Overview, story.Body, story.id as ThisID ,story.newsletterid, story.name as storyName";
                ThisSQL2 = ThisSQL2 + " from newsletterissues nl, newsletterissuepages story";
                ThisSQL2 = ThisSQL2 + " Where nl.active=1 and (NL.ID = story.newsletterid) ";
                ThisSQL2 = ThisSQL2 + " and nl.NewsletterID=" + NewsletterID + " "; // 01/13/2017 Search only in the same NewsletterID
                if (monthSelected != 0) {
                    ThisSQL2 = ThisSQL2 + " and month(nl.publishdate) = " + monthSelected;
                }
                if (yearSelected != 0) {
                    ThisSQL2 = ThisSQL2 + " and year(nl.publishdate) = " + yearSelected;
                }
                if (!string.IsNullOrEmpty(SearchKeywords)) {
                    ThisSQL2 = ThisSQL2 + " and ((story.Body like '%" + SearchKeywords + "%' )or (story.name  like '%" + SearchKeywords + "%') or (story.Overview  like '%" + SearchKeywords + "%'))";
                }
                ThisSQL2 = ThisSQL2 + "  ORDER BY PublishDate DESC";
                // 
                // Call cs.OpenSQL(ThisSQL2, "", RecordsPerPage, PageNumber)
                cs.OpenSQL(ThisSQL2, "");
                if (!cs.OK()) {
                    layout.load(newsArchiveListItemLayout);
                    layout.setClassInner("newsArchiveListCaption", "No results were found");
                    // Call layout.SetClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found"))
                    layout.setClassInner("newsArchiveListOverview", "");
                    Stream += layout.getHtml().Replace("?", $"?{cp.Utils.ModifyQueryString(refreshQueryString, Constants.RequestNameFormID, Constants.FormArchive.ToString(), true)}");  // layout.GetHtml()
                    // Stream &= cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found")
                } else {
                    layout.load(newsArchiveListItemLayout);
                    layout.setClassInner("newsArchiveListCaption", "Search results");
                    // Call layout.SetClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search Results Found", "Search results"))
                    layout.setClassInner("newsArchiveListOverview", "");
                    Stream += layout.getHtml().Replace("?", $"?{cp.Utils.ModifyQueryString(refreshQueryString, Constants.RequestNameFormID, Constants.FormArchive.ToString(), true)}"); // layout.GetHtml()
                    while (cs.OK() & RowCount < RecordsPerPage) {
                        storyName = cs.GetText("storyName");
                        storyOverview = cs.GetText("Overview");
                        storyBody = cs.GetText("body");
                        if (string.IsNullOrEmpty(storyOverview)) {
                            if (!NewsletterController.isBlank(cp, storyBody)) {
                                // if cs.GetBoolean("AllowReadMore") Then
                                storyOverview = storyBody;
                            } else {
                                storyOverview = cp.Content.GetCopy("Newsletter Article Access Denied", "You do not have access to this article");
                            }
                        }
                        qs = refreshQueryString;
                        // 01/12/2017 Dwayne request change the link to the full history
                        // qs = cp.Utils.ModifyQueryString(qs, "formid", FormCover.ToString())
                        qs = cp.Utils.ModifyQueryString(qs, "formid", Constants.FormStory.ToString());
                        qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameStoryId, cs.GetInteger("ThisID").ToString());
                        layout.load(newsArchiveListItemLayout);
                        layout.setClassInner("newsArchiveListCaption", storyName);
                        layout.setClassInner("newsArchiveListOverview", storyOverview);
                        if (layout.getHtml().Contains("?")) {
                            // cp.Utils.AppendLog("Test2.log", layout.getHtml())
                        }
                        Stream += layout.getHtml().Replace("href=\"?\"", $"href=\"?{qs}\"");
                        cs.GoNext();
                        RowCount = RowCount + 1;
                    }
                }
                //
                // 01/13/2017 comment pagination
                // 
                // If FileCount <> 0 Then
                // GoToPage = ""
                // Do While PageCount <= NumberofPages
                // qs = refreshQueryString
                // qs = cp.Utils.ModifyQueryString(qs, RequestNameButtonValue, FormButtonViewArchives)
                // qs = cp.Utils.ModifyQueryString(qs, RequestNamePageNumber, PageCount.ToString())
                // qs = cp.Utils.ModifyQueryString(qs, RequestNameSearchKeywords, SearchKeywords)
                // GoToPage &= "<a href=""?" & qs & """>" & (PageCount) & "</a>"
                // PageCount = PageCount + 1
                // GoToPage &= "&nbsp;&nbsp;&nbsp;"
                // Loop
                // Call layout.Load(newsArchiveListItemLayout)
                // Call layout.SetClassInner("newsArchiveListCaption", GoToPage)
                // Call layout.SetClassInner("newsArchiveListOverview", "")
                // Stream &= layout.GetHtml()
                // End If
            }
            // 
            if (!BlockSearchForm) {
                // 
                // Display search form
                // 
                string searchForm = "";
                searchForm += "<h2>Archive Search</h2>";
                // searchForm &= cp.Content.GetCopy("Newsletter Search Copy", "<h2>Archive Search</h2>")
                searchForm += "<div>" + cp.Html.SelectContent(Constants.RequestNameIssueID, "", Constants.ContentNameNewsletterIssues, "active=1 and (Publishdate<" + cp.Db.EncodeSQLDate(DateTime.Now) + ")AND(NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ")") + "</div>";
                // searchForm &= "<div>" & cp.Html.SelectContent(RequestNameIssueID, "", ContentNameNewsletterIssues, "(Publishdate<" & cp.Db.EncodeSQLDate(Now) & ")AND(NewsletterID=" & cp.Db.EncodeSQLNumber(NewsletterID) & ")") & " " & cp.Html.Button(FormButtonViewNewsLetter) & "</div>"
                searchForm += "<div>&nbsp;</div>";
                searchForm += "<div>keyword search<br>";
                searchForm += cp.Html.InputText(Constants.RequestNameSearchKeywords) + "</div>";
                MonthString = "";
                MonthString += "Month <select size=\"1\" name=\"" + Constants.RequestNameMonthSelectd + "\">";
                MonthString += "<option selected>Month</option>";
                for (MonthCounter = 1; MonthCounter <= 12; MonthCounter++) {
                    MonthString += "<option ";
                    MonthString += $"value=\"{MonthCounter}\">{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(MonthCounter)}</option>";
                }
                MonthString += "</select> ";
                //
                YearString = "";
                YearString += $"Year <select size=\"1\" name=\"{Constants.RequestNameYearSelected}\">";
                YearString += "<option selected>Year</option>";
                var loopTo = DateTime.Now.Year;
                for (YearCounter = DateTime.Now.Year - YearsWanted; YearCounter <= loopTo; YearCounter++) {
                    YearString += "<option ";
                    YearString += $"value=\"{YearCounter}\">{YearCounter}</option>";
                }
                YearString += "</select>";
                searchForm += "<div>&nbsp;</div>";
                searchForm += $"<div>{MonthString}&nbsp;&nbsp;&nbsp;{YearString}&nbsp;&nbsp;&nbsp;&nbsp;{cp.Html.Button("button", Constants.FormButtonViewArchives)}</div>";
                searchForm += "<div>&nbsp;</div>";
                searchForm += cp.Html.Hidden(Constants.RequestNameFormID, Constants.FormArchive.ToString());
                searchForm = cp.Html.Form(searchForm);
                //
                layout.load(newsArchiveListItemLayout);
                layout.setClassInner("newsArchiveListCaption", "");
                layout.setClassInner("newsArchiveListOverview", searchForm);
                Stream += layout.getHtml();
            }
            //
            //
            GetArchiveItemListRet = Stream;
            return GetArchiveItemListRet;
        }
        // 
        // 
        internal string GetSearchItemList(CPBaseClass cp, NewsletterController cn, string ButtonValue, int issueId, string refreshQueryString, string newsArchiveListItemLayout) {
            string GetSearchItemListRet = default;
            // 
            var layout = new BlockClass();
            int recordTop;
            int RecordsPerPage;
            int archiveIssuesToDisplay;
            var cs = cp.CSNew();
            var NewsletterID = default(int);
            int monthSelected;
            int yearSelected;
            string SearchKeywords;
            // 
            string link = "";
            string Stream = "";
            string ThisSQL;
            string ThisSQL2 = "";
            string MonthString;
            string YearString;
            int MonthCounter;
            int YearCounter;
            string storyName;
            var NumberofPages = default(int);
            int PageCount;
            string sql2;
            var FileCount = default(int);
            string storyOverview;
            var RowCount = default(int);
            int YearsWanted;
            var BlockSearchForm = default(bool);
            string qs;
            int PageNumber;
            DateTime issueDate;
            string issueDateFormatted = "";
            string GoToPage = "";
            string storyBody = "";
            // 
            // -- move to view
            monthSelected = cp.Doc.GetInteger(Constants.RequestNameMonthSelectd);
            yearSelected = cp.Doc.GetInteger(Constants.RequestNameYearSelected);
            SearchKeywords = cp.Doc.GetText(Constants.RequestNameSearchKeywords);
            recordTop = cp.Doc.GetInteger(Constants.RequestNameRecordTop);
            // 
            // todo -- these are now in the settings model
            var newsletter = Contensive.Models.Db.DbBaseModel.create<NewsletterModel>(cp, NewsletterID);
            archiveIssuesToDisplay = newsletter.archiveIssuesToDisplay;
            RecordsPerPage = 10;
            // 
            PageNumber = cp.Doc.GetInteger(Constants.RequestNamePageNumber);
            if (PageNumber == 0) {
                PageNumber = 1;
            }
            // 
            YearsWanted = cp.Utils.EncodeInteger(cp.Site.GetText("Newsletter years wanted", "1"));
            if (YearsWanted < 1) {
                YearsWanted = 1;
            }
            // 
            if (archiveIssuesToDisplay == 0) {
                archiveIssuesToDisplay = 6;
            }
            // 
            // Get Total Archive count
            // 
            PageCount = 1;
            sql2 = " select count(story.id) as count";
            sql2 = sql2 + " from newsletterissues nl, newsletterissuepages story";
            sql2 = sql2 + " Where nl.active=1 and (NL.ID = story.newsletterid)";
            sql2 = sql2 + " AND (NL.NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ")";
            if (monthSelected != 0) {
                ThisSQL2 = ThisSQL2 + " and month(nl.publishdate) = " + monthSelected;
            }
            if (yearSelected != 0) {
                ThisSQL2 = ThisSQL2 + " and year(nl.publishdate) = " + yearSelected;
            }
            if (!string.IsNullOrEmpty(SearchKeywords)) {
                sql2 = sql2 + " and ((story.Body like '%" + SearchKeywords + "%' )or (story.name  like '%" + SearchKeywords + "%') or (story.Overview  like '%" + SearchKeywords + "%'))";
            }
            if (cs.OpenSQL(sql2)) {
                FileCount = cs.GetInteger("count");
                NumberofPages = (int)Math.Round(FileCount / (double)RecordsPerPage);
                if (NumberofPages != (int)Math.Floor((double)NumberofPages)) {
                    NumberofPages = NumberofPages + 1;
                    NumberofPages = (int)Math.Floor((double)NumberofPages);
                }
                if (NumberofPages == 0) {
                    NumberofPages = 1;
                }
            }
            cs.Close();
            // 
            // Colors = "#ffffff"
            // 
            // 
            if ((ButtonValue ?? "") != Constants.FormButtonViewNewsLetter & (ButtonValue ?? "") != Constants.FormButtonViewArchives) {
                // 
                // List a page of archive issues
                // 
                if (monthSelected == 0 & yearSelected == 0) {
                    // stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                    // 
                    // ThisSQL = " SELECT  TOP 6 * From NewsletterIssues WHERE (PublishDate < { fn NOW() }) AND (ID <> " & IssueID & ") AND (NewsletterID=" & cp.db.encodesqlNumber(NewsletterID) & ") ORDER BY PublishDate DESC"
                    ThisSQL = " SELECT  TOP " + archiveIssuesToDisplay + " * From NewsletterIssues WHERE active=1 and (PublishDate < { fn NOW() }) AND (ID <> " + issueId + ") AND (NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ") ORDER BY PublishDate DESC";
                    // 
                    cs.OpenSQL(ThisSQL);
                    if (cs.OK()) {
                        while (cs.OK()) {
                            layout.load(newsArchiveListItemLayout);
                            issueDate = cs.GetDate("PublishDate");
                            if (issueDate != DateTime.MinValue) {
                                issueDateFormatted = $"{CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(issueDate.Month)} {issueDate.Day}, {issueDate.Year}";
                            }
                            link = refreshQueryString;
                            link = cp.Utils.ModifyQueryString(link, Constants.RequestNameIssueID, cs.GetInteger("ID").ToString());
                            layout.setClassInner("newsArchiveListCaption", cs.GetText("Name"));
                            layout.setClassInner("newsArchiveListOverview", cp.Utils.EncodeContentForWeb(cs.GetText("Overview")));
                            Stream += layout.getHtml().Replace("?", $"?{link}");
                            cs.GoNext();
                        }
                    } else {
                        BlockSearchForm = true;
                        layout.load(newsArchiveListItemLayout);
                        layout.setClassInner("newsArchiveListCaption", "<span class=\"ccError\">" + cp.Site.GetText(Constants.SitePropertyNoNewsletterArchives, "There are currently no archived issues.") + "</span>");
                        layout.setClassInner("newsArchiveListOverview", "");
                        Stream += layout.getHtml();
                    }
                    cs.Close();
                }
            }
            if ((ButtonValue ?? "") == Constants.FormButtonViewArchives) {
                // 
                // List search results of archive issues
                // 
                // stream &=  "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=5>"
                ThisSQL2 = " select NL.id, nl.name, nl.publishdate, story.AllowReadMore, story.Overview, story.Body, story.id as ThisID ,story.newsletterid, story.name as storyName";
                ThisSQL2 = ThisSQL2 + " from newsletterissues nl, newsletterissuepages story";
                ThisSQL2 = ThisSQL2 + " Where nl.active=1 and (NL.ID = story.newsletterid)";
                if (monthSelected != 0) {
                    ThisSQL2 = ThisSQL2 + " and month(nl.publishdate) = " + monthSelected;
                }
                if (yearSelected != 0) {
                    ThisSQL2 = ThisSQL2 + " and year(nl.publishdate) = " + yearSelected;
                }
                if (!string.IsNullOrEmpty(SearchKeywords)) {
                    ThisSQL2 = ThisSQL2 + " and ((story.Body like '%" + SearchKeywords + "%' )or (story.name  like '%" + SearchKeywords + "%') or (story.Overview  like '%" + SearchKeywords + "%'))";
                }
                ThisSQL2 = ThisSQL2 + "  ORDER BY PublishDate DESC";
                // 
                cs.OpenSQL(ThisSQL2, "", RecordsPerPage, PageNumber);
                if (!cs.OK()) {
                    layout.load(newsArchiveListItemLayout);
                    layout.setClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found"));
                    layout.setClassInner("newsArchiveListOverview", "");
                    Stream += layout.getHtml();
                    // Stream &= cp.Content.GetCopy("Newsletter Search No Results Found", "No results were found")
                } else {
                    layout.load(newsArchiveListItemLayout);
                    layout.setClassInner("newsArchiveListCaption", cp.Content.GetCopy("Newsletter Search Results Found", "Search results"));
                    layout.setClassInner("newsArchiveListOverview", "");
                    Stream += layout.getHtml();
                    while (cs.OK() & RowCount < RecordsPerPage) {
                        storyName = cs.GetText("storyName");
                        storyOverview = cs.GetText("Overview");
                        storyBody = cs.GetText("body");
                        if (string.IsNullOrEmpty(storyOverview)) {
                            if (!NewsletterController.isBlank(cp, storyBody)) {
                                // if cs.GetBoolean("AllowReadMore") Then
                                storyOverview = storyBody;
                            } else {
                                storyOverview = cp.Content.GetCopy("Newsletter Article Access Denied", "You do not have access to this article");
                            }
                        }
                        qs = refreshQueryString;
                        qs = cp.Utils.ModifyQueryString(qs, "formid", "400");
                        qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameStoryId, cs.GetInteger("ThisID").ToString());
                        layout.load(newsArchiveListItemLayout);
                        layout.setClassInner("newsArchiveListCaption", storyName);
                        layout.setClassInner("newsArchiveListOverview", storyOverview);
                        Stream += layout.getHtml().Replace("?", $"?{qs}");
                        cs.GoNext();
                        RowCount = RowCount + 1;
                    }
                }
                //
                if (FileCount != 0) {
                    GoToPage = "";
                    while (PageCount <= NumberofPages) {
                        qs = refreshQueryString;
                        qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameButtonValue, Constants.FormButtonViewArchives);
                        qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNamePageNumber, PageCount.ToString());
                        qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameSearchKeywords, SearchKeywords);
                        GoToPage += "<a href=\"?" + qs + "\">" + PageCount + "</a>";
                        PageCount = PageCount + 1;
                        GoToPage += "&nbsp;&nbsp;&nbsp;";
                    }
                    layout.load(newsArchiveListItemLayout);
                    layout.setClassInner("newsArchiveListCaption", GoToPage);
                    layout.setClassInner("newsArchiveListOverview", "");
                    Stream += layout.getHtml();
                }
            }
            // 
            if (!BlockSearchForm) {
                // 
                // Display search form
                // 
                string searchForm = "";
                searchForm += cp.Content.GetCopy("Newsletter Search Copy", "<h2>Archive Search</h2>");
                searchForm += "<div>" + cp.Html.SelectContent(Constants.RequestNameIssueID, "", Constants.ContentNameNewsletterIssues, "active=1 and (Publishdate<" + cp.Db.EncodeSQLDate(DateTime.Now) + ")AND(NewsletterID=" + cp.Db.EncodeSQLNumber(NewsletterID) + ")") + " " + cp.Html.Button(Constants.FormButtonViewNewsLetter) + "</div>";
                searchForm += "<div>&nbsp;</div>";
                searchForm += "<div>keyword search<br>";
                searchForm += cp.Html.InputText(Constants.RequestNameSearchKeywords) + "</div>";
                MonthString = "";
                MonthString += "Month <select size=\"1\" name=\"" + Constants.RequestNameMonthSelectd + "\">";
                MonthString += "<option selected>Month</option>";
                for (MonthCounter = 1; MonthCounter <= 12; MonthCounter++) {
                    MonthString += "<option ";
                    MonthString += $"value=\"{MonthCounter}\">{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(MonthCounter)}</option>";
                }
                MonthString += "</select> ";
                //
                YearString = "";
                YearString += $"Year <select size=\"1\" name=\"{Constants.RequestNameYearSelected}\">";
                YearString += "<option selected>Year</option>";
                var loopTo = DateTime.Now.Year;
                for (YearCounter = DateTime.Now.Year - YearsWanted; YearCounter <= loopTo; YearCounter++) {
                    YearString += "<option ";
                    YearString += "value=\"" + YearCounter + "\">" + YearCounter + "</option>";
                }
                YearString += "</select>";
                searchForm += "<div>&nbsp;</div>";
                searchForm += "<div>" + MonthString + "&nbsp;&nbsp;&nbsp;" + YearString + "&nbsp;&nbsp;&nbsp;&nbsp;" + cp.Html.Button(Constants.FormButtonViewArchives) + "</div>";
                searchForm += cp.Html.Hidden(Constants.RequestNameFormID, Constants.FormArchive.ToString());
                searchForm += cp.Html.Form(searchForm);
                // 
                layout.load(newsArchiveListItemLayout);
                layout.setClassInner("newsArchiveListCaption", "");
                layout.setClassInner("newsArchiveListOverview", searchForm);
                Stream += layout.getHtml();
            }
            // 
            // 
            GetSearchItemListRet = Stream;
            return GetSearchItemListRet;
        }
        // 
        private string GetFormRow(string Innards) {
            string GetFormRowRet = default;
            string Stream = "";
            // 
            Stream += "<TR>";
            Stream += "<TD colspan=2 width=\"60%\">" + Innards + "</TD>";
            Stream += "</TR>";
            GetFormRowRet = Stream;
            return GetFormRowRet;
        }
        // 
        private string GetSpacer(int Height = 1, int Width = 1) {
            string GetSpacerRet = default;
            // On Error GoTo ErrorTrap
            // 
            string Stream;
            // 
            Stream = "<img src=\"/ccLib/images/spacer.gif\" width=\"" + Width + "\" height=\"" + Height + "\">";
            // 
            GetSpacerRet = Stream;
            return GetSpacerRet;
            // 
            // Exit Function
            // ErrorTrap:
            // Call HandleError("LeftSideNavigation", "GetSpacer")
        }
        // '
        // Private Function GetArticleAccess(cp As CPBaseClass, ArticleID As Integer, isManager As Boolean, Optional GivenGroupID As Integer = 0) As Boolean
        // 'On Error GoTo ErrorTrap
        // '
        // Dim cs As CPCSBaseClass = cp.CSNew()
        // Dim AccessFlag As Boolean
        // Dim ThisTest As String
        // '
        // If GivenGroupID <> 0 Then
        // Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
        // If Not cs.OK() Then
        // GetArticleAccess = True
        // Else
        // Do While cs.OK()
        // If cs.GetInteger("GroupID") = GivenGroupID Then
        // GetArticleAccess = True
        // End If
        // Call cs.GoNext()
        // Loop
        // End If
        // Call cs.Close()
        // Else
        // If Not isManager Then
        // Call cs.Open(ContentNameNewsLetterGroupRules, "NewsletterPageID=" & ArticleID, , , , , "GroupID")
        // If Not cs.OK() Then
        // GetArticleAccess = True
        // Else
        // Do While cs.OK()
        // ThisTest = cs.GetText("GroupID")
        // '
        // '
        // If ThisTest <> "" Then
        // If cp.User.IsInGroup(ThisTest) Then
        // GetArticleAccess = True
        // End If
        // End If
        // Call cs.GoNext()
        // Loop
        // End If
        // Call cs.Close()
        // Else
        // GetArticleAccess = True
        // End If
        // End If
        // '
        // 'Exit Function
        // 'ErrorTrap:
        // 'Call HandleError(cp, ex, "GetArticleAccess")
        // End Function
        // 
        // Private Function GetIssuePublishDate(ByVal cp As CPBaseClass, ByVal IssueID As Integer) As String
        // 'On Error GoTo ErrorTrap
        // '
        // Dim cs As CPCSBaseClass = cp.CSNew()
        // Dim IssueDate As String
        // Dim Stream As String = ""
        // '
        // cs.Open(ContentNameNewsletterIssues, "ID=" & IssueID, , , "PublishDate")
        // If cs.OK Then
        // IssueDate = cs.GetDate("PublishDate")
        // If IsDate(IssueDate) Then
        // Stream = MonthName(Month(IssueDate), True) & " " & Day(IssueDate) & ", " & Year(IssueDate)
        // End If
        // End If
        // Call cs.Close()
        // '
        // '
        // GetIssuePublishDate = Stream
        // '
        // End Function
        // '
        internal string GetCoverContent(CPBaseClass cp, int IssueID, int storyId, string refreshQueryString, int formid, string newsCoverStoryItem, string newsCoverCategoryItem, bool isEditing, ref string return_Sponsor, ref DateTime return_publishDate, ref string return_tagLine) {
            string returnHtmlItemList = "";
            try {
                // 
                var layout = new BlockClass();
                var cs = cp.CSNew();
                string Criteria;
                string MainSQL;
                string CategoryName;
                var RecordCount = default(int);
                var cn = new NewsletterController();
                int CategoryID;
                string qs;
                string cover;
                //
                Constants.openRecord(cp, ref cs, "Newsletter Issues", IssueID);
                if (cs.OK()) {
                    cover = cs.GetText("Cover");
                    return_Sponsor = cs.GetText("sponsor");
                    return_tagLine = cs.GetText("tagLine");
                    return_publishDate = GenericController.encodeMinDate(cs.GetDate("publishDate"));
                    if (cover.Length > 50) {
                        returnHtmlItemList = GetCoverStoryItemLayout(cp, newsCoverStoryItem, "", "", "", cover, "", "", "", "", isEditing, cs.GetEditLink());
                    }
                }
                cs.Close();
                // 
                if (storyId != 0) {
                    Criteria = "";
                    MainSQL = "" + " select p.categoryId,c.name as CategoryName" + " from NewsletterIssuePages p" + " left join NewsletterIssueCategories c on c.id=p.categoryId" + " where (p.ID=" + cp.Db.EncodeSQLNumber(storyId) + ")" + "";




                    // call cs.open(ContentNameNewsletterStories, Criteria, "SortOrder,DateAdded")
                } else {
                    //
                    MainSQL = "SELECT DISTINCT NIC.ID AS CategoryID, NIR.SortOrder, NIC.Name AS CategoryName";
                    MainSQL = MainSQL + " FROM NewsletterIssueCategories NIC, NewsletterIssueCategoryRules NIR";
                    MainSQL = MainSQL + " Where (NIC.ID = NIR.CategoryID)";
                    MainSQL = MainSQL + " AND (NIR.NewsletterIssueID=" + IssueID + ")";
                    MainSQL = MainSQL + " AND (NIC.Active<>0)";
                    MainSQL = MainSQL + " AND (NIR.Active<>0)";
                    MainSQL = MainSQL + " ORDER BY NIR.SortOrder";
                    // 
                    // Call cp.Site.TestPoint("MainSQL: " & MainSQL)
                    // Call cs.OpenSQL(  MainSQL)
                    // 
                }
                cp.Site.TestPoint("MainSQL: " + MainSQL);
                cs.OpenSQL(MainSQL);
                // 
                if (cs.OK()) {
                    while (cs.OK()) {
                        CategoryID = cs.GetInteger("CategoryID");
                        using (var CS2 = cp.CSNew()) {
                            if (CS2.Open(Constants.ContentNameNewsletterStories, "(CategoryID=" + CategoryID + ") AND (NewsletterID=" + IssueID + ")", "SortOrder,id")) {
                                // 
                                // there are stories under this topic, wrap in div to allow a story indent
                                layout.load(newsCoverCategoryItem);
                                CategoryName = cs.GetText("CategoryName");
                                if (isEditing & RecordCount != 0) {
                                    qs = refreshQueryString;
                                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameIssueID, IssueID.ToString());
                                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameSortUp, CategoryID.ToString());
                                    CategoryName += "&nbsp;<a href=\"?" + qs + "\"><span style=\"font-family:helvetica,arial,san-serif;font-weight:Normal;font-size:13px;text-decoration:none;\">[Move Up]</span></a> ";
                                }
                                layout.setClassInner("newsCoverCategoryItem", CategoryName);
                                returnHtmlItemList += layout.getHtml();
                                // 
                                while (CS2.OK()) {
                                    returnHtmlItemList += GetCoverStoryItem(cp, CS2, formid, refreshQueryString, newsCoverStoryItem, isEditing, CS2.GetEditLink());
                                    CS2.GoNext();
                                }
                            }
                            CS2.Close();
                        }
                        cs.GoNext();
                        RecordCount = RecordCount + 1;
                    }
                }
                // 
                cs.Close();
                // 
                Criteria = "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" + IssueID + ")";
                cs.Open(Constants.ContentNameNewsletterStories, Criteria, "SortOrder,DateAdded");
                if (cs.OK()) {
                    // Caption = cp.Site.GetText("Newsletter Caption Other Stories", "")
                    // If Caption <> "" Then
                    // Stream &= vbCrLf & "<div class=""NewsletterTopic"">" & Caption & "</div>"
                    // End If
                    while (cs.OK()) {
                        returnHtmlItemList += GetCoverStoryItem(cp, cs, formid, refreshQueryString, newsCoverStoryItem, isEditing, cs.GetEditLink());
                        cs.GoNext();
                    }
                }
                cs.Close();
                // 
                if (isEditing) {
                    layout.load(newsCoverStoryItem);
                    layout.setClassInner("newsCoverListCaption", cp.Content.GetAddLink(Constants.ContentNameNewsletterStories, "Newsletterid=" + IssueID, false, cp.User.IsEditing()));
                    layout.setClassInner("newsCoverListOverview", "");
                    layout.setClassInner("newsCoverListReadMore", "");
                    layout.setClassInner("infographicBox", "");
                    returnHtmlItemList += layout.getHtml();
                }
            } catch (Exception ex) {
                handleError(cp, ex, "GetNewsletterBodyOverview");
            }
            return returnHtmlItemList;
        }
        // '

        // Private Function GetUnrelatedStories(ByVal cp As CPBaseClass, ByVal IssuePageID As Integer, ByVal IssueID As Integer, ByVal formId As Integer, ByVal refreshQueryString As String, ByVal newsCoverStoryItem As String) As String
        // Dim returnHtml As String = ""
        // Try

        // 'On Error GoTo ErrorTrap
        // '
        // Dim Criteria As String
        // Dim cs As CPCSBaseClass = cp.CSNew()
        // Dim Caption As String
        // '
        // If IssuePageID = 0 Then
        // Criteria = "((CategoryID is Null) OR (CategoryID=0)) AND (NewsletterID=" & IssueID & ")"
        // Call cs.Open(ContentNameNewsletterStories, Criteria, "SortOrder,DateAdded")
        // If cs.OK() Then
        // 'Caption = cp.Site.GetText("Newsletter Caption Other Stories", "")
        // 'If Caption <> "" Then
        // '    Stream &= vbCrLf & "<div class=""NewsletterTopic"">" & Caption & "</div>"
        // 'End If
        // Do While cs.OK()
        // returnHtml &= GetStoryOverview(cp, cs, formId, IssuePageID, refreshQueryString, newsCoverStoryItem)
        // Call cs.GoNext()
        // Loop
        // End If
        // Call cs.Close()
        // End If
        // Catch ex As Exception
        // Call handleError(cp, ex, "GetUnrelatedStories")
        // End Try
        // Return returnHtml
        // End Function
        // 
        private string GetCoverStoryItem(CPBaseClass cp, CPCSBaseClass CSStories, int formId, string refreshQueryString, string newsCoverStoryItem, bool isEditing, string editLink) {
            string returnhtml = "";
            try {
                // 
                int StoryID;
                string StoryAccessString;
                var cn = new NewsletterController();
                string storyBookmark;
                string caption = "";
                string readMoreLink = "";
                string overview = "";
                string storyBody = "";
                string coverInfographicthumbnail = "";
                string coverInfographic = "";
                string coverInfographicUrl = "";
                // 
                StoryID = CSStories.GetInteger("ID");
                coverInfographicthumbnail = CSStories.GetText("coverInfographicthumbnail");
                coverInfographic = CSStories.GetText("coverInfographic");
                coverInfographicUrl = CSStories.GetText("coverInfographicUrl");
                storyBookmark = "story" + StoryID;
                // 
                StoryAccessString = NewsletterController.GetArticleAccessString(cp, StoryID);
                // 
                if (formId != Constants.FormEmail) {
                    caption += CSStories.GetEditLink();
                }
                caption = "<span id=\"" + storyBookmark + "\">" + CSStories.GetText("Name") + "</span>";
                overview += cp.Utils.EncodeContentForWeb(CSStories.GetText("Overview"));
                storyBody = CSStories.GetText("body");
                if (!NewsletterController.isBlank(cp, storyBody)) {
                    readMoreLink = refreshQueryString;
                    readMoreLink = cp.Utils.ModifyQueryString(readMoreLink, Constants.RequestNameStoryId, StoryID.ToString());
                    readMoreLink = cp.Utils.ModifyQueryString(readMoreLink, Constants.RequestNameFormID, Constants.FormStory.ToString());
                }
                returnhtml = GetCoverStoryItemLayout(cp, newsCoverStoryItem, StoryAccessString, storyBookmark, caption, overview, readMoreLink, coverInfographicthumbnail, coverInfographic, coverInfographicUrl, isEditing, editLink);
            } catch (Exception ex) {
                handleError(cp, ex, "getStoryOverview");
            }
            return returnhtml;
        }
        // 
        // ====================================================================================================
        /// <summary>
        /// Populate an instance of the cover item template
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="newsCoverStoryItem"></param>
        /// <param name="StoryAccessString"></param>
        /// <param name="storyBookmark"></param>
        /// <param name="caption"></param>
        /// <param name="overview"></param>
        /// <param name="readMoreLink"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private string GetCoverStoryItemLayout(CPBaseClass cp, string newsCoverStoryItem, string StoryAccessString, string storyBookmark, string caption, string overview, string readMoreLink, string coverinfographicThumbnail, string coverinfographic, string coverInfographicUrl, bool isEditing, string editLink) {
            string returnhtml = "";
            try {
                // 
                var layout = new BlockClass();
                var cn = new NewsletterController();
                string readMore;
                string img = "";
                // 

                layout.load(newsCoverStoryItem);
                // 
                if (string.IsNullOrEmpty(coverinfographicThumbnail)) {
                    // 
                    // no infographic
                    // 
                    layout.setClassOuter("infographicBox", "");
                } else {
                    coverinfographicThumbnail = Uri.EscapeUriString(coverinfographicThumbnail);
                    img = "<img src=\"" + cp.Http.CdnFilePathPrefix + coverinfographicThumbnail + "\" alt=\"View the infographic\" class=\"banner\" width=\"100%\">";
                    if (string.IsNullOrEmpty(coverinfographic)) {
                        // 
                        // no image
                        // 
                        if (string.IsNullOrEmpty(coverInfographicUrl)) {
                            layout.setClassInner("infographImage", img);
                        } else {
                            if (coverInfographicUrl.IndexOf("://") < 0) {
                                coverInfographicUrl = "http://" + coverInfographicUrl;
                            }
                            coverInfographicUrl = Uri.EscapeUriString(coverInfographicUrl);
                            layout.setClassInner("infographImage", "<a href=\"" + coverInfographicUrl + "\" target=\"_blank\">" + img + "</a>");
                        }
                    } else {
                        // 
                        // linked thumbnail
                        // 
                        coverinfographic = Uri.EscapeUriString(coverinfographic);
                        layout.setClassInner("infographImage", "<a href=\"" + cp.Http.CdnFilePathPrefix + coverinfographic + "\" target=\"_blank\">" + img + "</a>");
                    }
                }
                if (string.IsNullOrEmpty(coverinfographic)) {
                    layout.setClassOuter("infographLink", "");
                } else {
                    layout.setClassInner("infographLink", "<a href=\"" + cp.Http.CdnFilePathPrefix + coverinfographic + "\" target=\"_blank\">View the infographic online.</a>");
                }
                if (!string.IsNullOrEmpty(StoryAccessString)) {
                    layout.prepend("<AC type=\"AGGREGATEFUNCTION\" name=\"block text\" querystring=\"allowgroups=" + StoryAccessString + "\">");
                }
                if (string.IsNullOrEmpty(caption)) {
                    layout.setClassOuter("newsCoverListCaption", "");
                } else {
                    layout.setClassInner("newsCoverListCaption", caption);
                }
                if (string.IsNullOrEmpty(overview)) {
                    layout.setClassOuter("newsCoverListOverview", "");
                } else {
                    layout.setClassInner("newsCoverListOverview", overview);
                }

                if (string.IsNullOrEmpty(readMoreLink)) {
                    layout.setClassOuter("newsCoverListReadMore", "");
                } else {
                    readMore = layout.getClassInner("newsCoverListReadMore");
                    readMore = readMore.Replace("?", "?" + readMoreLink);
                    readMore = readMore.Replace("#", "?" + readMoreLink);
                    layout.setClassInner("newsCoverListReadMore", readMore);
                }
                if (!string.IsNullOrEmpty(StoryAccessString)) {
                    layout.append("<AC type=\"AGGREGATEFUNCTION\" name=\"block text end\" >");
                }
                // 
                returnhtml = layout.getHtml();
                if (isEditing) {
                    returnhtml = $"{editLink}<div class=\"ccEditWrapper\">" + returnhtml + "</div>";
                }
            } catch (Exception ex) {
                handleError(cp, ex, "getStoryOverview");
            }
            return returnhtml;
        }
        // 
        internal string GetStory(CPBaseClass cp, NewsletterController cn, int storyId, int IssueID, string refreshQueryString, string newsBody, bool isEditing) {
            string returnHtml = "";
            try {
                var cs = cp.CSNew();
                var CSIssue = cp.CSNew();
                bool rssChange;
                var PublishDate = default(DateTime);
                int Pos;
                string Copy;
                string Link;
                string PrinterIcon;
                string EmailIcon;
                string storyName;
                string storyOverview;
                string storyBody;
                string qs = "";
                var layout = new BlockClass();
                // 
                layout.load(newsBody);
                // 
                PrinterIcon = "<img border=0 src=/ccLib/images/IconPrint.gif>";
                EmailIcon = "<img border=0 src=/ccLib/images/IconEmail.gif>";
                // 
                if (storyId == 0) {
                    layout.setClassOuter("newsBodyCaption", "");
                    layout.setClassInner("newsBodyStory", "<span class=\"ccError\">The requested story is currently unavailable.</span>");
                } else {
                    cs.Open(Constants.ContentNameNewsletterStories, "ID=" + storyId);
                    if (cs.OK()) {
                        storyName = cs.GetText("name");
                        if (isEditing) {
                            storyName = cs.GetEditLink() + storyName;
                        }
                        storyBody = cs.GetText("body");
                        storyOverview = cs.GetText("Overview");
                        if (string.IsNullOrEmpty(storyBody)) {
                            storyBody = storyOverview;
                        }
                        IssueID = cs.GetInteger("newsletterId");
                        // 
                        returnHtml += cs.GetEditLink();
                        if (cs.GetBoolean("AllowPrinterPage")) {
                            qs = cp.Doc.RefreshQueryString;
                            qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameStoryId, storyId.ToString());
                            qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormStory.ToString());
                            qs = cp.Utils.ModifyQueryString(qs, "ccIPage", "l6d09a10sP");
                            returnHtml += "<div class=\"PrintIcon\"><a target=_blank href=\"?" + qs + "\">" + PrinterIcon + "</a>&nbsp;<a target=_blank href=\"" + qs + "\"><nobr>Printer Version</nobr></a></div>";
                        }
                        if (cs.GetBoolean("AllowEmailPage")) {
                            Link = $"mailto:?SUBJECT={cp.Site.GetText("Email link subject", $"A link to the {cp.Site.DomainPrimary} newsletter")}&amp;BODY=http://{cp.Site.DomainPrimary}{cp.Request.Page}{refreshQueryString.Replace("&", "%26")}{Constants.RequestNameStoryId}={storyId}%26{Constants.RequestNameFormID}={Constants.FormStory}";
                            returnHtml += "<div class=\"EmailIcon\"><a target=_blank href=\"?" + Link + "\">" + EmailIcon + "</a>&nbsp;<a target=_blank href=\"" + Link + "\"><nobr>Email this page</nobr></a></div>";
                        }
                        layout.setClassInner("newsBodyCaption", storyName);
                        layout.setClassInner("newsBodyStory", storyBody);
                        // 
                        // update RSS fields if empty
                        // 
                        if (!isEditing) {
                            rssChange = false;
                            if (IssueID != 0) {
                                if (NewsletterController.encodeMinDate(cs.GetDate("RSSDatePublish")) == DateTime.MinValue) {
                                    CSIssue.Open(Constants.ContentNameNewsletterIssues, "id=" + cp.Db.EncodeSQLNumber(IssueID));
                                    if (CSIssue.OK()) {
                                        PublishDate = CSIssue.GetDate("publishDate");
                                    }
                                    CSIssue.Close();
                                    if (NewsletterController.encodeMinDate(PublishDate) != DateTime.MinValue) {
                                        rssChange = true;
                                        cs.SetField("RSSDatePublish", PublishDate.ToString());
                                    }
                                }
                            }
                            // 
                            if (!string.IsNullOrEmpty(storyName) & string.IsNullOrEmpty(cs.GetText("RSSTitle"))) {
                                rssChange = true;
                                cs.SetField("RSSTitle", cs.GetText("name"));
                            }
                            // 
                            if (!string.IsNullOrEmpty(storyOverview) & string.IsNullOrEmpty(cs.GetText("RSSDescription"))) {
                                rssChange = true;
                                Copy = cp.Utils.ConvertHTML2Text(storyOverview);
                                cs.SetField("RSSDescription", Copy);
                            }
                            // 
                            if (string.IsNullOrEmpty(cs.GetText("RSSLink"))) {
                                Link = cp.Request.Link;
                                if (Link.IndexOf(cp.Site.GetText("adminUrl"), StringComparison.OrdinalIgnoreCase) < 0) {
                                    Pos = Link.IndexOf("?", StringComparison.Ordinal);
                                    if (Pos >= 0) {
                                        Link = Link.Substring(0, Pos);
                                    }
                                    qs = refreshQueryString;
                                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameStoryId, storyId.ToString());
                                    qs = cp.Utils.ModifyQueryString(qs, Constants.RequestNameFormID, Constants.FormStory.ToString());
                                    qs = cp.Utils.ModifyQueryString(qs, "method", "");
                                    rssChange = true;
                                    cs.SetField("RSSLink", Link + "?" + qs);
                                }
                            }
                            if (rssChange) {
                                RssFeedController.updateRSSFeed(cp);
                            }
                        }
                    }
                    cs.Close();
                }
                // 
                return layout.getHtml();
            } catch (Exception ex) {
                handleError(cp, ex, "getNewsletterBodyDetails");
                throw;
            }
        }
        // 
        private string template(int x) {
            string returnHtml = "";
            try {

            } catch (Exception) {
                // Call handleError(cp, ex, "template")
            }
            return returnHtml;
        }
        // 
        // =====================================================================================
        // common report for this class
        // =====================================================================================
        // 
        private void handleError(CPBaseClass cp, Exception ex, string @method) {
            try {
                cp.Site.ErrorReport(ex, "Unexpected error in newsletterBodyClass." + @method);
            } catch (Exception) {
                //
                // stop anything thrown from cp errorReport
                //
            }
        }
    }
}