# alfred-outlook
Alfred outlook mail search
This is a Alfred Workflow for searching Microsoft Outlook Version 16 and later.
This workflow supports searching:
* Mail Subject
* Mail Sender
* Mail Preview Content

By default, it will search all three above togther, you can figure out which single item you want to search with below format:
- olk from:\[keyword\]
- olk title:\[keyword\]

It supports pagination.
Supports multiple keywords search:
- olk \[keyword1\] \[keyword2\] \[keyword3\] 
- olk title:\[keyword1\] \[keyword2\] \[keyword3\] 

You can use olkc to set some configuration for search:
* olkc pagesize \[number\]
* olkc folder
  * it will list all available folders in your outlook account, then select one as your default search target
* olkc profile
  * it will list all available profiles (account) in your outlook, then select one as your preferred
* olkc filter \[SQL like filter\]
  * allow to set a SQL like filter for Subject, with '%' support fuzzy matching, for example:
    * '%test' to filter all Subjects end with 'test'
    * 'test%' to filter all Subjects start with 'test'
    * '%test%' to filter all Subject contain 'test'
    * (Limitation: only allow one filter in current version 0.1.8)

Attachment Download (**Alfred V3** up only)
* After search result listed, an icon with 'clip' means this mail contains attachment, Press 'CTRL' and click, all attachments will be save to disk
* Default Path to save is '~/Downloads/outlook_attachments', you can change it to your preferred folder with changing Alfred Workflow Environment Variables:
  * Open 'Alfred Preferences' -> Click on tab 'Workflows' -> Choose 'Outlook Search' in sidebar
  * Click the [x] in the top right of workflow to show the Workflow Environment Variables panel
  * There is a variable named 'olk_attachment_path' with default value '~/Downloads/outlook_attachments', change it to your own folder
  * All attachments will be save to a sub-folder named with your mail subject with 'Folder Allowed Characters'
  * After attachments downloaded, saved path will be in clipboard

Powerful Pagination:
* 'Next Page' if there's more pages available.
* If you are using **Alfred V3**, Press 'CTRL' on 'Next Page' item as modifier, then it behave as 'Previous Page'

Compose a new mail
* olknew \[optional: email address\]

Search Contact (Person)
* olkp \[keyword\]
   * it will search contacts by display name and email address
   * for found results
     * you can select to open contact pane in Outlook
     * If you are using **Alfred V3**, Press 'CTRL' on selected contact for composing a new mail to contact in Outlook

Download built version here:

https://github.com/xeric/alfred-outlook/releases
