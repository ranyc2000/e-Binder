# e-Binder
Dead simple automated document binder.

This project creates a web page that serves as an electronic binder to organize and browse a document collection. It can be run without a web server. There are no server-side processes at all. All the action is generated within the client browser, via Javascript and jQuery.

The main web page file is a template (container) that never changes. What changes (at design time) is the data pulled into the main web page (via Javascript), from a data file. Excel is used as a tool (at design time, not run time), to create the data file.  The data file is automatically generated by an Excel (VBA) macro, triggered by the button you will see in Excel called "Generate." Excel is used here as strictly an authoring/maintenance tool. It is not needed to run or use the project.

The documents to be viewed in the binder live wherever you designate, on the local machine or a remote machine. By default, the document repo is in a folder within the project folder ( binder/images ).

To use the provided example as is, copy (pull) the entire project (including subfolders) to your drive and double-click index.html.
You will see a sample binder with 30+ documents on "valves" of various types.

To adapt this to your own content, replace the contents of the IMAGES folder with your own document collection.

You will also need to generate a new data file. 

Capture a directory of the documents:

In Windows Command shell (Win-R CMD), C:\BINDER\IMAGES\> DIR /B /S \*.\* > listing.txt

 ... and paste the results into the Full Path column (F) of the spreadsheet (replacing what was there). If you have global edits (say you want to enter a path prefix to each line, or the like) it is easier to paste the listing into your favorite editor (e.g., MS Word), make the edits and then copy and past into Col. F of the spreadsheet (and you can then discard the editor file).
 
Fill in the other columns Date, Doc Type, Title, Summary, as desired. For starters you can just reuse the file spec data (Col. F) for both the Title and the Summary and revise later.

Go to the Data Tab and fill in the Binder Name, sOutDir-PC (or Mac) (which is the full path to your binder folder) and the number of items you want to display per page.
Go back to the Documents tab and Save. Then press Generate. (On a Mac, the operating system may require you to approve a write permission for this.)
Then check your handiwork by launching index.html.

That's it!

Now for some information

This is a tiny templated system. The user need not edit any HTML, Javascript or Visual Basic to deploy a new binder. All that needs to be added or edited is the data itself, for the index. It requires a web browser and a place to put the content and control files, which can be a local or remote file system without even requiring a web server (though it will also work under a web server).

The data is entered in Excel. An Excel macro, generateData, spits out a Javascript file, gendata.js that captures all the index data. (The operative VBA macro is embedded in the Excel file; the separate file generateData.vba is provided for documentation purposes, to make it easier to see this code, if you are interested.)

The provided file index.html reads in the Javascript from gendata.js and recreates the data structure of the Excel file in running Javascript. index.html uses the loaded contents of gendata.js as a data source during run time (it does not actually need runtime access to the Excel file itself). index.html also reads in some jQuery libraries to get some more UI capabilities to jazz up the presentation.

Styling comes from the file igen.css. There are copies of the html and css in tabs of the spreadsheet, also for archival purposes (these tabs don't actually do anything). The file gendata.js is generated (from the Excel) and specific to each project. 

You never need to edit index.html, igen.css, the Excel macro, or anything else, other than the data in the Excel file.

All that has to change to make a new project are the Documents and Data tabs of the Excel, and your actual document collection.

You can put the IMAGES folder whereever you want and name it whatever you want - so long as Col. F has the proper paths to your files.
(../ goes up a folder level, etc., and you should be able to rech any file in your file system.) You can also use unlimited nested directories for your documents (even unrelated directories), again, so long as the paths in Col. F reflect the paths to each document.

This works on both Windows and Mac. You therefore have lots of latitude in how you lay out your data and these control files.

The Excel macro will convert forward slashes to backslashes in the document path for use in Windows. Mac doesn't care.

The Data tab in the preadsheet mentions search capability. This version does not have it.

The site you create does not require a server. It runs directly off the file system. You do not need to be online. 
So if you put this on a laptop and have no net connection, it will work just fine. Nor do you even need Excel on the machine you are running the viewer on. (Excel is out of the picture once you hit "Generate.")

Things may change a bit in the page display depending what viewers you have installed in your browser. It is not too picky about the browser either.

Note that there is a lot of room in each item for a Summary. The examples here are short, but they can be much longer. You can write an entire analysis of each document, and it will appear at the bottom of the left pane when you are viewing the document. This is a major way we improve over the native file viewer in your File Manager. We can also be selective about which files we include in the index. Also, the date column gives you a place to put the *real* document date, rather than something picked up by the file system.

You can use Excel before you Generate to sort the spreadsheet by date, to get a chronological listing. Undated docs will all sort at one end of the list, so you may want to put in approximate dates (even if that is something like Jan. 1 of some year) rather than leave them blank.

The VBA commands in the Excel macro work on both PC and Mac - you can use Office on the PC or the Mac - either will work.

(The following is not a big deal, but note that Excel carries forward sort of a "high water mark" for the data it has stored (Used Range), which can result in blank lines that Excel treats as part of the spreadsheet. If the document collecton you are cataloging has fewer items than the one you are replacing, you should prefereably reset the Used Range in the spreadsheet before you Generate. Hit Ctrl-End to find what Excel thinks is the last cell. Highlight all rows from the *real* last row of actual data, through the last row that Excel *sees*, and right-click-Delete those blank rows. Save and reopen and test again with Ctrl-End to make sure the empty rows are really gone.)

Is this simple or what? Hope you like it.

Comments and enhancements welcome!
