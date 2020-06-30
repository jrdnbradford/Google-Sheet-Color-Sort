# Google-Sheet-Color-Sort
Google Sheet-bound script that assists with sorting Google Sheet rows by selected column's background fill color.

Note: Google recently added the ability to sort Sheets by background fill color. See their [help article](https://support.google.com/docs/answer/3540681) for details.

## Usage
### Add the Script to your Google Sheet
In your Google Sheet select **Tools** -> **Script editor**. Remove the default code, paste in the contents of this [script](googleSheetColorSort.gs), save, and enter a project title of your choice. Close the script editor and go back to your Sheet.

### Authorize the Application
Refresh the Sheet, click **Google Sheet Color Sort** in the menu bar, and click any of the submenu items. Then authorize the application using your Google account (this must only be done once). Once authorized you must click **Google Sheet Color Sort** again to see your sorting options.

### Select a Sorting Column
Select a column that contains the background fill colors you wish to sort on by clicking and highlighting any cell within that column.

### Choose a Sorting Method
Once you have highlighted a cell within your column of choice, use the **Google Sheet Color Sort** menu button to select a sorting option.

#### Sort Rows by Color
Use the **Sort Rows by Color** submenu item to automatically sort the entire Google Sheet by each row's background fill color in your selected column. 

#### Add Sorting Column
Use the **Add Sorting Column** submenu item to add a column that contains each row's background fill color as a hexadecimal color code that can be used to manually sort sections of the Google Sheet by selecting the relevant rows, going to **Data** -> **Sort Range**, and selecting the column with the values. 

After selecting a submenu item, specify whether the data in your Sheet has a header row in the window that appears.

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