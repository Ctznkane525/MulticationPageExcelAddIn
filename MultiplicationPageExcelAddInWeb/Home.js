(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            /*
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }
            */

            config();

            $('#button-generate').text("Configure");
            $('#button-generate').click(configure);
                
            //loadSampleData();

           
        });
    };

    function config() {
        $.getJSON("config.json").then((config) => {
            $("#numRows").val(config.rows);
            $("#numCols").val(config.cols);
            $("#numMax").val(config.max);
            $("#numMin").val(config.min);
        });
    }

    function randomInteger(min, max) {
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }

    function componentToHex(c) {
        var hex = c.toString(16);
        return hex.length == 1 ? "0" + hex : hex;
    }

    function rgbToHex(r, g, b) {
        return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
    }

    function configure() {

        let numRows = parseInt($("#numRows").val());
        let numCols = parseInt($("#numCols").val());
        let numMax = parseInt($("#numMax").val());
        let numMin = parseInt($("#numMin").val());

        let getRandomValues = (leadingText) => {
            let subItems = [];
            let colNum = 0;
            for (colNum = 0; colNum < numCols; colNum++) {
                subItems.push(leadingText + randomInteger(numMin, numMax));
            }
            return subItems;
        }

        let getBlankValues = (leadingText) => {
            let subItems = [];
            let colNum = 0;
            for (colNum = 0; colNum < numCols; colNum++) {
                subItems.push("");
            }
            return subItems;
        }

        let items = [];
        let rowNum = 0;
        for (rowNum = 0; rowNum < numRows; rowNum++) {

            let subItems = [];
            items.push(getRandomValues(" "));
            items.push(getRandomValues("*"));

            items.push(getBlankValues(""));
            items.push(getBlankValues(""));
        }


        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            let workbook = ctx.workbook;

            var sheet = workbook.worksheets.getActiveWorksheet();

            var cells = sheet.getRangeByIndexes(0, 0, items.length, numCols);
            cells.values = items;
            cells.format.font.color = rgbToHex(86, 50, 168);
            cells.format.horizontalAlignment = "Right";
            cells.format.font.size = 16;
            cells.format.columnWidth = 72;
 
            let i = 0;
            for (i = 2; i < items.length; i+=4)
            {
                var affectedRow = sheet.getRangeByIndexes(i, 0, 1, numCols);
                affectedRow.format.borders.getItem('EdgeTop').style = 'Continuous';      
            }

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
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
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
