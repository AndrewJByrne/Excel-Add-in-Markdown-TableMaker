/// <reference path="/Scripts/FabricUI/MessageBanner.js" />
/// <reference path="./markdown-table-maker.js"/>

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
                    // I have now loaded all the data I need to produce markdown for the selected range.
                    var markdownString = MarkdownTableMaker.makeMarkdownTable(cells);
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
