# alfred-outlook
Alfred outlook mail search

<table style="padding: 5px; widt: 50%;">
    <tr>
        <td>
        olk &#124; olkp &#124; olknew &#124; olkc
        </td>
    </tr>
</table>


----------------------------------------

This is an Alfred Workflow for searching Microsoft Outlook Version 16 and later.

## Command olk - search email

This workflow supports searching:
- Mail Subject
- Mail Sender
- Mail Preview Content

By default, it will search all three above togther, you can figure out which single item you want to search with below format:

    olk from:{keyword}
or

    olk title:{keyword}

Supports multiple keywords search:

    olk {keyword1} {keyword2} {keyword3}

or

    olk title:{keyword1} {keyword2} {keyword3}

### This workflow search result supports Powerful Pagination:
> * 'Next Page' if there's more pages available.
> * If you are using **Alfred V3**, Press 'CTRL' on 'Next Page' item as modifier, then it behave as 'Previous Page'

## Command olkc - configuration

You can use olkc to set some configurations for search:

    olkc pagesize {number}
>It will change your search result list size for every page

    olkc folder
>It will list all available folders in your outlook account, then select one as your default search target

    olkc profile
>It will list all available profiles (account) in your outlook, then select one as your preferred

## Command olknew - Compose a new mail

    olknew {optional: email address}


>if email address is not assigned, this will new an Outlook compose mail window without receiver, otherwise email address will be set as receiver.

## Command olkp - Search Contact (Person)

    olkp {keyword}


 >it will search contacts by display name and email address.
 >> * for found results, you can select to open contact pane in Outlook.
 >> * If you are using **Alfred V3**, Press 'CTRL' on selected contact for composing a new mail to contact in Outlook

## Download built version here:

https://github.com/xeric/alfred-outlook/releases
