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
                onGenerateClick);

            // Add a click event handler for the copy button.
            $('#copy-button').click(
                onCopyClick);
        });
    }

   
    function onGenerateClick() {
        Excel
        .run(loadSelectedCells)
        .then(generateMarkdown)
        .then(showMarkdownString)
        .catch(showError);
    }

    function loadSelectedCells(ctx) {
            
        // Create a proxy object for the selected range and load some properties
        var selectedRange; 
        var cells = [];

        selectedRange = ctx.workbook.getSelectedRange().load("rowCount, columnCount");
        return ctx.sync().then(function () {
            for (var r = 0; r < selectedRange.rowCount; r++) {
                var col = [];
                for (var c = 0; c < selectedRange.columnCount; c++) {
                    col.push(selectedRange.getCell(r, c).load("format/font/*, values"));
                }
                cells.push(col);
            }
            return cells;
        });
    }

    function generateMarkdown(cells) {
        for (var r = 0; r < cells.length; r++) {
            cells[r] = cells[r].map(function (cell) {
                return {
                    bold: cell.format.font.bold,
                    italic: cell.format.font.italic,
                    value: cell.values[0][0],
                };
            });
        }

        var markdownString = MarkdownTableMaker.makeMarkdownTable(cells)

        return markdownString;
    }

    function showMarkdownString(tableMarkdownAsString) {
        $("#markdown-result").text(tableMarkdownAsString);
    }

    function onCopyClick() {

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
    function showError(error) {
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
        .catch(showError);
    }
})();
