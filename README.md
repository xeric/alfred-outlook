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
