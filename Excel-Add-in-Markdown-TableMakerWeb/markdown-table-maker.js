// markdown-table-maker.js
var MarkdownTableMaker = (function() {
  
    // expose to public
    return {
        // public name   : name of internal function
        makeMarkdownTable: makeMarkdownTable,
    }

    // all private
  
    /**
     * Takes in an array of Excel cells and returns a string that represents this range in 
     * table Markdown.
     */
    function makeMarkdownTable(cells) {
        var markdownString = "";

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

        return markdownString;
    }

    // Create markdown for the given cell usign the value as well as
    // formatting info. 
    function markdownize(cell) {
        var value = detectUrl(cell.value);

        if (cell.bold) {
            value = "**" + value + "**";
        }

        if (cell.italic) {
            value = "_" + value + "_";
        }
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

        // Regex to detect a URL
        function isUrl(text) {
            return (typeof (text) === 'string') && /[(https?)|(file)]:\/\/.+$/.test(text);
        }

        // Regex to detect and image file name
        function isImage(text) {
            return (typeof (text) === 'string') && /.+\.(jpeg|jpg|gif|png)$/.test(text);
        }
    }

})();

