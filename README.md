# Google-Sheet-Color-Sort
Google Sheet-bound script that assists with sorting Google Sheet rows by background fill color.

## Using the Script
In your Google Sheet select **Tools** -> **Script editor**. Remove the default code, paste in the contents of this [script](googleSheetColorSort.gs), save, and enter a project title of your choice. Close the script editor and go back to your Sheet.

Refresh the Sheet, click **Google Sheet Color Sort** in the menu bar, and click any of the submenu items. Then authorize the application using your Google account (this must only be done once). Once authorized you must click **Google Sheet Color Sort** again to see your sorting options.

Use the **Sort Rows by Color** submenu item to automatically sort the entire Google Sheet by each row's background fill color. 

Use the **Add Sorting Column** submenu item to add a column with each row's background fill color as a hexadecimal color code that can be used to manually sort sections of the Google Sheet by selecting the relevant rows, going to **Data** -> **Sort Range**, and selecting the column with the values. 

After selecting a menu item, specify whether the data in your Sheet has a header row in the window that appears.

## Recommended OAuth Scope
```json
{
    "oauthScopes": [
        "https://www.googleapis.com/auth/spreadsheets.currentonly"
    ]
}
```

## Authors
**Jordan Bradford** - GitHub: [jrdnbradford](https://github.com/jrdnbradford)

## License
This script is licensed under the MIT license. See [LICENSE.txt](LICENSE.txt) for details.