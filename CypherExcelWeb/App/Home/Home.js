/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#execute').click(function() {
                var query = $('#query').val();
                var url = $('#url').val();
                executeQuery(query, url);
            });
        });
    };

    function executeQuery(query, url) {
        app.closeNotification();
        $('.disable-while-executing').prop('disabled', true);
        var cypherEndpoint = url + '/db/data/cypher';
        $.ajax({
                type: 'POST',
                url: cypherEndpoint,
                accepts: 'application/json',
                dataType: 'json',
                data: { 'query': query }
            })
            .success(function (result) {
                var tableData = new Office.TableData(result.data, result.columns);
                pushTableToPage(tableData);
            })
            .fail(function(result) {
                app.showNotification(result.statusText);
            })
            .always(function() {
                $('.disable-while-executing').prop('disabled', false);
            });
    }

    function pushTableToPage(tableData) {
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