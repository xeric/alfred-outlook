# alfred-outlook
***Alfred outlook mail search***

This is an Alfred Workflow for searching Microsoft Outlook Version 16 and later.
```
olk | olkp | olknew | olkc
```
----------------------------------------

## Command olk - Search email

This workflow supports searching:
- Mail Subject
- Mail Sender
- Mail Preview Content

By default, it will search all three above together, you can figure out which single item you want to search with below format:

    olk from:{keyword}
<h>
    
    olk title:{keyword}

Supports multiple keywords search:

    olk {keyword1} {keyword2} {keyword3}
<h>

    olk title:{keyword1} {keyword2} {keyword3}

It also support get last {n} unread mail quickly:

    olk recent:10

or just get today's unread mail:

    olk recent:today

You can also apply some simple keywords filter on it:

    olk recent:30 {keyword1} {keyword2} 

## Attachment Download (**Alfred V3** and up only)
After search result listed, an icon with <img src="https://raw.githubusercontent.com/xeric/alfred-outlook/master/attachment.png" width="20" height="20"> means this mail contains attachment, Press 'CTRL' and click, all attachments will be save to disk

Default Path to save is '~/Downloads/outlook_attachments', you can change it to your preferred folder with changing Alfred Workflow Environment Variables:
  * Open 'Alfred Preferences' -> Click on tab 'Workflows' -> Choose 'Outlook Search' in sidebar
  * Click the [x] in the top right of workflow to show the Workflow Environment Variables panel
  * There is a variable named 'olk_attachment_path' with default value '~/Downloads/outlook_attachments', change it to your own folder
  * All attachments will be save to a sub-folder named with your mail subject with 'Folder Allowed Characters'
  * After attachments downloaded, saved path will be in clipboard

### This workflow search result supports Powerful Pagination:
> * 'Next Page' if there's more pages available.
> * If you are using **Alfred V3**, Press 'CTRL' on 'Next Page' item as modifier, then it behave as 'Previous Page'

## Command olkc - Configuration

You can use olkc to set some configurations for search:

    olkc pagesize {number}
>It will change your search result list size for every page

    olkc folder
>It will list all available folders in your outlook account, then select one as your default search target

    olkc profile
>It will list all available profiles (account) in your outlook, then select one as your preferred

    olkc filter {SQL like filter}
>It allows to set a SQL like filter for Subject, with '%' support fuzzy matching, for example:
>   * '%test' to filter all Subjects end with 'test'
>   * 'test%' to filter all Subjects start with 'test'
>   * '%test%' to filter all Subject contain 'test'
>   * (Limitation: only allow one filter in current version)

## Command olknew - Compose a new mail

    olknew {optional: email address}

There are two mode of compose a new mail in plugin, activate Outlook App window mode and inline mode (If you are using **Alfred V3**).

If you are using **Alfred V3**,
to use inline mode, after you type a email address after olknew command, you can choose: 'Send a Direct Quick Mesaage', and you can type a message after hint 'Message: ' then send it without activating Outlook App window.

> if you are first time using this feature, and when you send a mail through inline mode, you will get a popup warning said:

> A script is attempting to send a message. Some scripts can contain viruses or otherwise be harmful to your computer, so it's important to verify that the script was created by a trustworthy source.
Do you want to send the message?

when you checked 'Don't show this message again', this warning dialog won't show in next time you send a inline mail.


## Command olkp - Search Contact (Person)

    olkp {keyword}


 >it will search contacts by display name and email address.
 > * for found results, you can select to open contact pane in Outlook.
 > * If you are using **Alfred V3**, Press 'CTRL' on selected contact for composing a new mail to contact in Outlook

## Download built version here:

https://github.com/xeric/alfred-outlook/releases
