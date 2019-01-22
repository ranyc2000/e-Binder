# e-Binder
Automated document binder

This project creates a web page that serves as an electronic binder for a document collection.

The web page is automatically generated by an Excel macro, triggered by the button you will see in Excel called "Generate."

To view this as is, copy the entire project (including subfolders) to your drive and double-click index.html.
You will see a sample binder with 30+ documents on "valves" of various types.
To adapt this to your own content, replace the contents of the IMAGES folder with your own document collection.
Then capture a directory of the documents:

In Windows, C:\BINDER\IMAGES\> DIR /B /S *.* > listing.txt

 ... and paste the results into the Full Path column (F).
 
Fill in the other columns Date, Doc Type, Title, Summary, as desired. For starters you can just reuse the file spec data (Col. F) for the Title and Summary and revise later.

Go to the Data Tab and fill in the Binder Name, sOutDir-PC (or Mac) (which is the full path to your binder folder) and the number of items you want to display per page.
Go back to the Documents tabe and Save. Then press Generate.
Then check your handiwork by launching index.html.
That's it!

Now for some information

The Excel macro, generateData, spits out a Javascript file, gendata.js that captures all the index data.
The file index.html reads in the Javascript from gendata.js and recreates the data structure of the Excel file in running Javascript.
The file index.html also reads in some jQuery libraries to get some more UI capabilities.
Styling comes from igen.css.

You never have to touch gendata.js, the css file, the Excel macro, or index.html - though you are invited to diddle!

All that has to change to make a new project are the Documents and Data tabs of the excel, and your actual document collection.

You can put the IMAGES folder whereever you want and name it whatever you want - so long as Col. F has the proper paths to your files.
You can use unlimited nested directories for your documents, again, so long as the paths in Col. F reflect them.
This works on both Windows and Mac.

The Excel macro will convert forward slashes to backslashes in the document path for use in Windows. Mac doesn't care.

The Data tab in the preadsheet mentions search capability. This version does not have it.

The site you create does not require a server. It runs directly off the file system. You do not need to be online. 
So if you put this on a laptop and have no net connection, it will work just fine.

Things may change a bit in the display depending what viewers you have installed.

Note that there is a lot of room in each item for a Summary. You can write an entire analysis and it will appear at the bottom of the left pane when you are viewing the document.

Comments and changes welcome!
