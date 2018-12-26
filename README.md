# google-sheet-h2oathome

This project is to manage door-to-door sales. It is especially done for H20 at home and to manage his account through google sheet.

It will create 3 sheets:

* Workshop: this sheet will contains all workshop done with the sales amount, all participants and how many they need to pay.
* Data: this sheet will contains all movement done in the account and then, each total amount of each workshop (Sales amount, Amount received by participants + commission)
* Gran total: this sheet is a pivot table to display benefits of the h2o activity

## Installation

To have the new menu h2o in your google sheet. You need to create a new sheet where you want to put all your activity.

* Click `Tools -> script editor`
* Copy the Code.gs file in the default created file
* Click `File -> New -> HTML file` and name it `NewWorkshop.html`
* Copy the "NewWorkshop.html" in this file
* Save all files and you should see the `H2O at home` menu appeared

## Using

The first time you use the plugin, you need to click on the menu `h2O at home -> Prepare sheet...`

This will create all sheet that the plugin need to work.

### New workshop

When you have a new workshop, you can click on the related menu and a popup will be displayed to get some information about the workshop.
When you click on the `Create` button, it will add rows in the **Workshop** sheet to put the detail of the workshop and the total line will be added in the **Data** sheet.

### Data

The data sheet will contains all data, amount that moves from your pro account. Then, you can add row without problem. Make `insert row before` the row 2. The sort is descending.