/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

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

            $("#template-description").text("Create table Markdown from the range of cells you select.");
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

            // Create a proxy object for the selected range and load its address and values properties
            var sourceRange = ctx.workbook.getSelectedRange().load("text, values, address, rowIndex, columnIndex, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {

                    if (sourceRange.rowCount < 2 || sourceRange.columnCount < 1) {
                        showNotification("No data selected", "Please select a range of data and try again.");
                    }
                    else {
                        // I am demonstrating two ways of generating the markdown string. This first method uses the 
                        // join method on an Array object. The second method uses brute-force, iterating over every element
                        // in the array. The first method also has fewer loops and no conditionals. On large ranges, this method
                        // out-performs the sencodn method by over a factor of 10. However, perf seems to be negligible for 
                        // small ranges and the second method offers more flexibility. 

                        console.time('Function #1');
                        // First row is the header row
                        markdownString = markdownString.concat('| ');
                        markdownString = markdownString.concat(sourceRange.values[0].join('| '));
                        markdownString = markdownString.concat('|\n');

                        // Add the header delimeter
                        markdownString = markdownString.concat('| ');
                        for (var cCount = 0; cCount < sourceRange.columnCount; cCount++) {
                            markdownString = markdownString.concat('---');
                            markdownString = markdownString.concat('| ');
                        }
                        markdownString = markdownString.concat('\n');

                        // Now at the rest of the rows
                        for (var i = 1; i < sourceRange.rowCount; i++) {
                            markdownString = markdownString.concat('| ');
                            markdownString = markdownString.concat(sourceRange.values[i].join('| '));
                            markdownString = markdownString.concat('\n');
                        }
                        console.timeEnd('Function #1');

                        //console.time('Function #2');
                        //markdownString = "";
                        //for (var i = 0; i < sourceRange.rowCount; i++) {

                        //    markdownString = markdownString.concat('| ');

                        //    for (var j = 0; j < sourceRange.columnCount; j++) {

                        //        markdownString = markdownString.concat(sourceRange.values[i][j]);
                        //        if (j <= sourceRange.columnCount - 1) {
                        //            markdownString = markdownString.concat('| ');
                        //        }

                        //        if (i == 0 && j == sourceRange.columnCount - 1) {
                        //            // This is the header row, so I need to add a row of 3-dash columns
                        //            markdownString = markdownString.concat('\n');
                        //            markdownString = markdownString.concat('| ');
                        //            for (var cCount = 0; cCount < sourceRange.columnCount; cCount++) {
                        //                markdownString = markdownString.concat('---');
                        //                markdownString = markdownString.concat('| ');
                        //            }

                        //        }

                        //    }
                        //    markdownString = markdownString.concat('\n');
                        //}
                        //console.timeEnd('Function #2');
                    }

                })
                   // Run the queued-up commands
                .then(ctx.sync)
                .then(function () {
                    if (markdownString.length > 0) {
                        showNotification("Table markdown generated!", "");

                        $("#markdown-result").text(markdownString);
                    }

                })
                .then(ctx.sync)
        })
        .catch(errorHandler);
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
