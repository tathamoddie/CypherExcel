/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayDataOrRedirect();
        });
    };

    // Checks if a binding exists, and either displays the visualization,
    //     or redirects to the Data Binding page.
    function displayDataOrRedirect() {
        Office.context.document.bindings.getByIdAsync(
            app.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var binding = result.value;
                    displayDataForBinding(binding);
                    // And bind a change-event handler to the binding:
                    binding.addHandlerAsync(
                        Office.EventType.BindingDataChanged,
                        function () {
                            displayDataForBinding(binding);
                        }
                    );
                } else {
                    window.location.href = '../DataBinding/DataBinding.html';
                }
            });
    }

    // Queries the binding for its data
    function displayDataForBinding(binding) {
        binding.getDataAsync(
            {
                coercionType: Office.CoercionType.Matrix,
                valueFormat: Office.ValueFormat.Unformatted,
                filterType: Office.FilterType.OnlyVisible
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    displayDataHelper(result.value);
                } else {
                    $('#data-display').html(
                        '<div class="notice">' +
                        '    <h2>Error fetching data!</h2>' +
                        '    <a href="../DataBinding/DataBinding.html">' +
                        '        <b>Bind to a different range?</b>' +
                        '    </a>' +
                        '</div>');
                }
            }
        );
    }

    // Displays data, once it has already been read off of the binding
    function displayDataHelper(data) {
        var rowCount = data.length;
        var columnCount = (data.length > 0) ? data[0].length : 0;
        if (!visualization.isValidRowAndColumnCount(rowCount, columnCount)) {
            $('#data-display').html(
                '<div class="notice">' +
                '    <h2>Not enough data!</h2>' +
                '    <p>The range must contain ' + visualization.rowAndColumnRequirementText + '.</p>' +
                '    <a href="../DataBinding/DataBinding.html">' +
                '        <b>Choose a different range?</b>' +
                '    </a>' +
                '</div>');
            return;
        }

        var $visualizationContent = visualization.createVisualization(data);

        $('#data-display').empty();
        $('#data-display').append($visualizationContent);
    }
})();