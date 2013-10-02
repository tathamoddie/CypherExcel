var visualization = (function () {
    "use strict";

    var visualization = {};

    // Sample data:
    visualization.sampleHeaders = [['Name', 'Grade']];
    visualization.sampleRows = [
        ['Ben', 79],
        ['Amy', 95],
        ['Jacob', 86],
        ['Ernie', 93]];

    // Data range validation:
    visualization.rowAndColumnRequirementText = '2 columns and at least 2 rows';
    visualization.isValidRowAndColumnCount = function (rowCount, columnCount) {
        return (rowCount > 1 && columnCount === 2);
    };

    // Creates a visualization, based on passed-in data:
    visualization.createVisualization = function (data) {
        var maxBarWidthInPixels = 200;

        var $table = $('<table class="visualization" />');
        var $headerRow = $('<tr />').appendTo($table);
        $('<th />').text(data[0][0]).appendTo($headerRow);
        $('<th />').text(data[0][1]).appendTo($headerRow);

        for (var i = 1; i < data.length; i++) {
            var $row = $('<tr />').appendTo($table);
            var $column1 = $('<td />').appendTo($row);
            var $column2 = $('<td />').appendTo($row);

            $column1.text(data[i][0]);
            var value = data[i][1];
            var width = (maxBarWidthInPixels * value / 100.0);
            var $visualizationBar = $('<div />').appendTo($column2);
            $visualizationBar.addClass('bar')
                .width(width)
                .text(value);
        }

        return $table;
    };

    return visualization;
})();