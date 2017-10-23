var app = angular.module('mainApp', []);
app.controller('mainController', function($scope) {
    var table_data = total_data;
    $scope.colData = [
        {name: "Tournament", width: 1},
        {name: "Date", width: 1},
        {name: "Infoslide", width: 1},
        {name: "Motion", width: 3},
        {name: "Round", width: 1},
        {name: "City", width: 1},
        {name: "Country", width: 1},
        {name: "Circuit", width: 1},
        {name: "CAs", width: 1},
        {name: "Topics", width: 1},
    ];
    $scope.filters = {
        Tournament: "",
        Date: "",
        Infoslide: "",
        Motion: "",
        Round: "",
        City: "",
        Country: "",
        Circuit: "",
        CAs: "",
        Topics: "",
    };
    page_size = 10;
    $scope.page = 1;
    $scope.maxPage = Math.ceil(table_data.length/page_size);
    $scope.table_contents = [];
    
    $scope.updatePagination = function(page) {
        page = page > $scope.maxPage ? $scope.maxPage : page;
        page = page < 1 ? 1 : page;
        $scope.page = page;
        start = ($scope.page-1)*page_size;
        end = $scope.page*page_size;
        end = end > table_data.length ? table_data.length : end;
        $scope.current_data = table_data.slice(start, end);
    }

    $scope.getTableStr = function(data) {
        switch (typeof(data)) {
            case "object":
                if (data instanceof Array) {
                    array_string = "";
                    delimiter = "";
                    for (e of data) {
                        array_string += delimiter + e;
                        delimiter = ", "
                    }
                    return array_string;
                }
            default:
                return data.toString();
        }
    }

    $scope.updateTableData = function() {
        var new_table = [];
        for (motion of total_data) {
            var failed_filter = false;
            for (header in $scope.filters) {
                motion_string = $scope.getTableStr(motion[header]).toLowerCase();
                filter_string = $scope.filters[header].toLowerCase();
                if (!motion_string.includes(filter_string)) {
                    failed_filter = true;
                }
            }
            if (!failed_filter) {
                new_table.push(motion);
            }
        }
        table_data = new_table;
        $scope.maxPage = Math.ceil(table_data.length/page_size);
        $scope.updatePagination(1);
    }

    $scope.updatePagination($scope.page);
});