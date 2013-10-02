/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#execute').click(function() {
                var query = $('#query').text();
                var url = $('#url').text();
                executeQuery(query, url);
            });
        });
    };

    function executeQuery(query, url) {
        var sampleHeaders = [['m','length(p)']];
        var sampleRows = [
            ['(5 {name:"Morpheus"})', 1],
            ['(4 {name:"Trinity"})', 2],
            ['(3 {name:"Cypher"})', 2],
            ['(2 {name:"Agent Smith"})', 3]
        ];

        var tableData = new Office.TableData(sampleRows, sampleHeaders);
        Office.context.document.setSelectedDataAsync(tableData,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    app.showNotification('Could not insert data',
                        'Please choose a different cell selection.');
                } else {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: app.bindingID },
                        function (asyncResult) {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                app.showNotification('Error binding data');
                            }
                        }
                    );
                }
            }
        );
    }

})();