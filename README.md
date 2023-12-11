# CRYPTOTOOLS_EXCEL
A library for importing ones balances, networth, staking, rewards, lending &amp; farming rates, dex volume &amp; fees, uniswap new pairs into Excel
## CRYPTOPRICE
https://user-images.githubusercontent.com/53000607/230736654-236c948e-1bb1-4070-bf93-98e24f812e3b.mov
## CRYPTOBALANCE
https://user-images.githubusercontent.com/53000607/230930296-559dda4f-1e77-4409-8156-d9e28e53f42f.mov
## CRYPTONETWORTH
https://user-images.githubusercontent.com/53000607/231311168-3832e9b7-4060-454f-b62c-31daac29f474.mov

## Step-by-Step Guide to Installing CryptoTools in Excel

This guide provides detailed instructions on how to install CryptoTools in Excel, including adding necessary modules and references.

### Adding the Developer Tab in Excel

1. Open Excel and go to the `File` tab.
2. Select `Options` to open the Excel Options dialog box.
3. In the dialog, choose `Customize Ribbon`.
4. In the right pane, check the box next to `Developer`.
5. Click `OK` to add the Developer tab to your Excel ribbon.

### Adding a New Module from the Developer Tab

1. Click on the `Developer` tab in the Excel ribbon.
2. Select `Visual Basic` to open the VBA editor.
3. In the VBA editor, right-click on `VBAProject (YourWorkbookName)` in the Project Explorer.
4. Choose `Insert`, then `Module` to add a new module.

### Installing CryptoTools.bas as the First Module

1. In the VBA editor with the new module selected, go to `File` > `Import File`.
2. Navigate to the location of your `CryptoTools.bas` file.
3. Select the file and click `Open` to import it.

### Installing JsonConverter.bas as the Second Module

1. Repeat the process of adding a new module.
2. Import `JsonConverter.bas` using the same method as for `CryptoTools.bas`.

### Adding Microsoft Scripting Runtime Reference

1. In the VBA editor, click on `Tools` > `References`.
2. Scroll and find `Microsoft Scripting Runtime`.
3. Check the box next to it.
4. Click `OK` to add the reference.

After following these steps, your Excel workbook will be set up with the necessary modules and references for CryptoTools.

