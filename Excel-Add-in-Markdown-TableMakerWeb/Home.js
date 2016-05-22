/// <reference path="/Scripts/FabricUI/MessageBanner.js" />
/// <reference path="./MarkdownTableMaker.js"/>

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#generate-button').click(
                    displaySelectedCells);
                return;
            }

            $("#template-description").text("Select a cell range and then tap the Generate button to produce the table Markdown.");
            $('#button-text').text("Generate!");
            $('#button-desc').text("Generates table Markdown for the selected range.");

            $('#copy-button-text').text("Copy");
            $('#copy-button-desc').text("Copies Markdown to clipboard");
                
            loadSampleData();

            // Add a click event handler for the generate button.
            $('#generate-button').click(
                generateTableMarkdown);

            // Add a click event handler for the copy button.
            $('#copy-button').click(
                copyToClipboard);
        });
    }

    function loadSampleData() {

        var values = [
                        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
                        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
                        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function generateTableMarkdown() {
        var markdownString = "";

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            
            // Create a proxy object for the selected range and load some properties
            var sourceRange = ctx.workbook.getSelectedRange().load("rowIndex, columnIndex, rowCount, columnCount");
            
            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {

                    if (sourceRange.rowCount < 1 || sourceRange.columnCount < 1) {
                        showNotification("No data selected", "Please select a range of data and try again.");
                    }
                    else {
                        console.time('Get Cells');
                        var rows = [];
                        for (var i = 0; i < sourceRange.rowCount; i++) {
                            var col = [];
                            for (var j = 0; j < sourceRange.columnCount; j++) {

                                // Create a proxy object for a 1-cell range and load its value and format
                                // Note: Because this is a range, I am retrieving a values array, even though
                                // it will contain just one value for this one cell. 
                                var cell = sourceRange.getCell(i, j).load(["values","format/*", "format/font"]);
                                col.push(cell);
                            }
                            rows.push(col);
                        }
                        console.timeEnd('Get Cells');
                        return rows;

                    }

                })
                   // Run the queued-up commands
                .then(ctx.sync)
                .then(function (cells) {

                    var markdownString = MarkdownTableMaker.makeMarkdownTable(cells);
                    if (markdownString.length > 0) {
                        showNotification("Table markdown generated!", "");

                        $("#markdown-result").text(markdownString);
                    }

                    // I have now loaded all the data I need to produce markdown for the selected range.

                    // I am demonstrating two ways of generating the markdown string. This first method uses the 
                    // join method on an Array object. The second method uses brute-force, iterating over every element
                    // in the array. The first method also has fewer loops and no conditionals. On large ranges, this method
                    // out-performs the second method by over a factor of 10. However, perf seems to be negligible for 
                    // small ranges and the second method offers more flexibility. 

                    console.time('Method #1');

                    // First row is the header row
                    markdownString = markdownString.concat('| ');
                    markdownString = markdownString.concat(cells[0].map(markdownize).join('| '));
                    markdownString = markdownString.concat('|\n');

                    // Add the header delimeter
                    markdownString = markdownString.concat('| ');
                    for (var cCount = 0; cCount < cells.length; cCount++) {
                        // Note: By adding colons to left and right of hyphens in the 
                        // header delimeter row, I am making all content center-align
                        markdownString = markdownString.concat(':---:');
                        markdownString = markdownString.concat('| ');
                    }
                    markdownString = markdownString.concat('\n');

                    // Now the rest of the rows
                    for (var i = 1; i < cells.length; i++) {
                        markdownString = markdownString.concat('| ');
                        markdownString = markdownString.concat(cells[i].map(markdownize).join('| '));
                        markdownString = markdownString.concat('|\n');
                    }
                    console.timeEnd('Method #1');

                    //console.time('Method #2');
                    //markdownString = "";
                    //for (var i = 0; i < cells.length; i++) {

                    //    markdownString = markdownString.concat('| ');

                    //    for (var j = 0; j < sourceRange.columnCount; j++) {

                    //        markdownString = markdownString.concat(markdownize(cells[i][j]));
                    //        if (j <= sourceRange.columnCount - 1) {
                    //            markdownString = markdownString.concat('| ');
                    //        }

                    //        if (i == 0 && j == sourceRange.columnCount - 1) {
                    //            // This is the header row, so I need to add a row of 3-dash columns
                    //            markdownString = markdownString.concat('\n');
                    //            markdownString = markdownString.concat('| ');
                    //            for (var cCount = 0; cCount < cells[0].length; cCount++) {
                    //                markdownString = markdownString.concat('---');
                    //                markdownString = markdownString.concat('| ');
                    //            }

                    //        }

                    //    }
                    //    markdownString = markdownString.concat('\n');
                    //}
                    //console.timeEnd('Method #2');

                    

                })
                .then(ctx.sync)
        })
        .catch(errorHandler);
    }

    // Create markdown for the given cell usign the value as well as
    // formatting info. 
    function markdownize(cell)
    {
        // I always get an array of values in this 1-cell range. 
        var value = cell.values[0][0];
        value = detectUrl(value);
        value = addSugar(value, cell.format);
        return value;
    }

    // Checks whether the value in a cell is a URL and generates the Markdown to 
    // represent it properly as a link. Also handles image URLs too. 
    function detectUrl(value) {
        var newValue = value;

        if (isUrl(value)) {
            var prefix = isImage(value) ? "!" : "";

            newValue = prefix + "[" + value + "](" + value + ")";
        }
        return newValue;
    }

    // Regex to detect a URL
    function isUrl(text) {
        return (typeof (text) === 'string') && /[(https?)|(file)]:\/\/.+$/.test(text);
    }

    // Regex to detect and image file name
    function isImage(text) {
        return (typeof (text) === 'string') && /.+\.(jpeg|jpg|gif|png)$/.test(text);
    }

    // Use the formatting info on the cell to add markup for a bold style 
    function addSugar(value, format) {
        if (format.bold) {
            value = "**" + value + "**";
        }

        if (format.italic) {
            value = "_" + value + "_";
        }

        return value;
    }

    function copyToClipboard() {

        // Make sure the text in markdown-result is selected
        $("#markdown-result").select();

        // Call Copy
        document.execCommand('Copy');
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
