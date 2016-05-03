# ExcelXY
Pronounced 'ex-cel-exy', why? because...

Author: D. Vincent West

Version: 2016.05.03

**Description:** An excel extension that adds convenient functionality for engineers, especially when working with XY datasets

**Acknowledgements** Most of this code was based on snippets I found around the internet from MVPs like Jon Peltier, Chip Pearson and others.

**Warning** This add-in uses a lot of VBA.  When macros are executed, the undo list is wiped out.  So, whenever you click a button on the ribbon, you cannot go back from the point.  Use with caution (save and backup often).

## Features

### Ribbon and Context Menu Integration
Many of the features are available on the ribbon, and also through integrated right-click context menus.  I don't have time to do a full write-up, but if you just start playing with it, you'll see some nice features that weren't there before

### Selecting by Datasets by Columns
Many of the Functionality is built in with the assumption that you want to do things with Columns of data.  Userforms are all (I think) built with the idea that you can select either a column, or simply one cell, and then the code will extrapolate downward to select the column of data that is there.  If the first cell does not contain a number, it is assumed to be a title, or a label. This makes it easier to do things with data.  Its not deeply feature full, just assuming you have simple columns of data that may or may not have a label.  If it doesn't make sense, just play with it.

### Charts:
* Click to Zoom functionality with AppEvents
	* Shift-click to set one rectangle corner, then again to set the other corner and the plot scales to that rectangle with a nice-autoscaling adjustment to account for excel's wacky scaling
* A simple auto-format option that create much nicer looking X-Y charts
* Nice Autoscaling (replaces Excel built-in autoscales)
* **Many of these options can be perfomed on multiple charts at once by selecting them with shift-click**
* Copy series data to the clipboard as a string
* Easy-add XY series data for faster modification of charts
* Export Charts to Image
* a few others, you'll have to play with it

### 1D Array Functions
* Calculate Derivatives and Integrals
	* Userforms for easy selection
	* Perform calculations on XY scatter charts

### Excel VBA Functions
I don't have time to document it here, but much of the code can be reused to build your own extensions

### Other
* Create a copy of the current workbook and open it
* Convert Formulas to Values for the entire workbook
* I'm sure there are more things to say, but I don't really have time to make it an exhaustive list right now.

### Contributing and Future Development
I make this add-in to speed up my workflow and also because its fun.  I try to make it as robust as I can, but it is probably a far cry from "professional" softare.  Things will break from time to time.  Going forward, I will ship updates.  If you want to contribute, let me know, and I will incorporate what I can.  I don't aim to make any money from this right now, its really just for fun.