/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#insert-sample-data').click(insertSampleData);
            $('#bind-to-existing-data').click(bindToExistingData);
        });
    };

    function insertSampleData() {
        var sampleData = new Office.TableData(
            visualization.sampleRows,
            visualization.sampleHeaders);
        Office.context.document.setSelectedDataAsync(sampleData,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    app.showNotification('Could not insert sample data',
                        'Please choose a different selection range.');
                } else {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: app.bindingID },
                        function (asyncResult) {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                app.showNotification('Error binding data');
                            } else {
                                window.location.href = '../Home/Home.html';
                            }
                        }
                    );
                }
            }
        );
    }

    function bindToExistingData() {
        Office.context.document.bindings.addFromSelectionAsync(
            Office.BindingType.Matrix,
            { id: app.bindingID },
            function (result) {
                var isValid = (result.status == Office.AsyncResultStatus.Succeeded) &&
                    visualization.isValidRowAndColumnCount(
                        result.value.rowCount, result.value.columnCount);
                if (isValid) {
                    window.location.href = '../Home/Home.html';
                } else {
                    app.showNotification('Invalid data selected',
                        'Please make a different selection, and ensure that you selected ' +
                        'a table or range with ' + visualization.rowAndColumnRequirementText);
                }
            }
        );
    }
})();