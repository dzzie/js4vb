// Classic JavaScript that will work
(function() {
    'use strict';
    
    function Calculator() {
        this.result = 0;
    }
    
    Calculator.prototype = {
        add: function(x) {
            this.result += x;
            return this;
        },
        subtract: function(x) {
            this.result -= x;
            return this;
        },
        multiply: function(x) {
            this.result *= x;
            return this;
        },
        getValue: function() {
            return this.result;
        }
    };
    
    var calc = new Calculator();
    calc.add(5).multiply(3).subtract(2);
    console.log(calc.getValue());
    
    // Object with nested functions
    var utils = {
        helpers: {
            format: function(str) {
                return str.toUpperCase();
            },
            validate: function(obj) {
                for (var key in obj) {
                    if (obj.hasOwnProperty(key)) {
                        if (typeof obj[key] === 'undefined') {
                            return false;
                        }
                    }
                }
                return true;
            }
        }
    };
})();