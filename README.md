# Excel-Add-in-Markdown-TableMaker

[Markdown](https://en.wikipedia.org/wiki/Markdown) is a lightweight markup language with plain text formatting syntax designed so that it can be converted to other formats. It is the language of choice when creating readme files on GitHub. Creating tables is not part of the code Markdown spec but they are supported by popular derivatives such as *Github-flavored Markdown (GFM)* and *Markdown Here*. In GFM, a table is specified using a series of dashes and pipes. Adam Pritchard's [Markdown Cheatsheet](https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet#tables)  does a great job of explaining how to create them. 

Excel's tabular style lends itself nicely to creating tables of data, but how do we convert to table Markdown for use in a Markdown (.md) file? This sample to the rescue! It takes a range of data you have selected and outputs the Markdown that is needed to represent that range as a table. It is written as an Excel Add-in and uses the Excel JavaScript APIs to load, iterate over and read a range of data from a spreadsheet.  

## Try it out

### Visual Studio version
1.  Copy the project to a local folder and open the Excel-Add-in-Markdown-TableMaker.sln in Visual Studio.

    > Make sure the project **Excel-Add-in-Markdown-TableMaker** is set as the startup project in the solution. To run the add-in in the Excel desktop client, set the **Start Action** to *Office Desktop Client*
    
    
2.  Press *F5* to build and deploy the sample add-in. Excel launches with a empty worksheet and a new command group is added to the ribbon.  
        
  ![](https://github.com/AndrewJByrne/Excel-Add-in-Markdown-TableMaker/blob/master/readme-images/launch.PNG)

3.  Tap the command labelled **TabMD** in the ribbon to open the task pane and populate the spreadsheet with test data. 

  ![](https://github.com/AndrewJByrne/Excel-Add-in-Markdown-TableMaker/blob/master/readme-images/open-tab.PNG)
  
4.  Select a range of cells in the spreadsheet and tap the **Generate!** button in the add-in.

  ![](https://github.com/AndrewJByrne/Excel-Add-in-Markdown-TableMaker/blob/master/readme-images/generate.PNG)
  
5.  The table markdown for this range is generated and displayed in the task pane's text field. 
6.  You can copy this markdown by selecting **Copy** and then paste it into a markdown (.md) file to preview in your favorite markdown editor. 

### Further work
To see a list of enhancements or to log an issue, visit the [issues](https://github.com/AndrewJByrne/Excel-Add-in-Markdown-TableMaker/issues) page of this repo. 


### Learn more

The Excel JavaScript APIs have much more to offer you as you develop add-ins. The following is a list of resources to help you learn more.  

* [Excel Add-ins programming overview](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
*  [Snippet Explorer for Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Excel Add-ins code samples](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
* [Excel Add-ins JavaScript API Reference](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
* [Build your first Excel Add-in](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
