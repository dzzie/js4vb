// Test potentially missing operators

// 1. Unsigned right shift
var a = -1;
var b = a >>> 1;  // Should give large positive number
print(">>> test: " + b);

// 2. Unsigned right shift assignment
var c = -1;
c >>>= 1;
print(">>>= test: " + c);

// 3. Exponentiation (ES2016 - probably not supported, that's OK)
// var d = 2 ** 3;  // Would be 8

// 4. Nullish coalescing (ES2020 - probably not supported)
// var e = null ?? "default";

// 5. Optional chaining (ES2020 - probably not supported)
// var f = obj?.prop?.nested;

// 6. Ternary operator
var g = (5 > 3) ? "yes" : "no";
print("Ternary: " + g);

// 7. typeof operator
var h = typeof "hello";
print("typeof: " + h);

// 8. void operator
var i = void 0;
print("void: " + i);

// 9. in operator
var obj = {foo: 1};
var j = "foo" in obj;
print("in operator: " + j);

// 10. instanceof operator  
var arr = [1,2,3];
var k = arr instanceof Array;
print("instanceof: " + k);

