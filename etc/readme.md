# Newsletter Template Guide

The newsletter was created in the time between "creating the html in the code" and mustache. The way it works is
* there is a template that can be selected in the newsletter record, from a table called "newsletter templates". You can use the search at the top of the admin site to find this table of templates.
* Currently it uses the wysiwyg editor. We know now that is a bad idea and if this is what we are working on I think the first thing is to make it a plain html code editor instead.
* These templates work using the class and id attributes in the html layout.
For example
* class newsHeaderMasthead -- code replaces the inner-html (html inside the tag) this with an uploaded file
* class newsFooterMasthead -- code replaces in the inner-html with a footer image

## Template Uses
One thing that makes the newsletter special is that same layout (newsletter template) is used for the website and for the email that it produces automatically. So all html elements have to be email compatible, and all the styles needed have to be included in the layout (no bootstrap for example)
There are 4 views created from the layout
* Cover Page -- what is emailed, and what is displayed when you go to the page
* Archive page -- shows a list of all the previous newsletters (I think)
* Story Page -- when someone clicks on a story on the cover page, this view is shown
* Search Page -- shows search results (dont remember where the search button is, but its there somewhere)

## Styles Controlling the Layout


* emailLinkToWeb -- only for email. The innerHtml is set to an anchor tag that takes the user to the online version of the Newsletter
* newsArchiveList -- the innerHtml is populated with a list created from the outerHtml of newsArchiveListItem
* newsArchiveListItem -- the item used to create newsArchiveList
* newsArchiveLink -- th code replaces the # sign in the innerHtml with a link to the archive page
* newsArchiveSearch -- The innerHtml is set to a search control that when used, directs them to the search results page
* newsBody -- the wrapper around the html used to create the story
* newsBodyCaption -- caption in the body
* newsBodyStory -- wrapper around the body of story within the body
* newsSearchListItem -- the item html used in the newsSearchList
* newsSearchList
* newsIssueCaption
* newsIssuePublishDate
* newsIssueSponsor
* newsletterTagLine
* adBannerItem -- not used
* newsArchive -- not used
* newsCover -- not used
