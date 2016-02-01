(function (scope) {
    scope.countrymerge = function (data, countries) {
        return data.regions.reduce(function (memo, region, i) {
            if (countries.indexOf(region) === -1) {
                memo.push(region);
            } else {
                for (var idx = region + 1; idx < (data.regions[i + 1] || data.names.length); idx++) {
                    memo.push(idx);
                }
            }

            return memo;
        }, []);
    };
})((typeof exports !== 'undefined' && exports) || window.Globalmigration || (window.Globalmigration = {}));
