// Test Function constructor
print("=== Testing Function Constructor ===");
// Simple function
var add = new Function('a', 'b', 'return a + b;');
print("add(5, 3) = " + add(5, 3));

// No parameters
var sayHello = new Function('return "Hello from dynamic function!";');
print(sayHello());

// Multiple parameters
var multiply = new Function('x', 'y', 'z', 'return x * y * z;');
print("multiply(2, 3, 4) = " + multiply(2, 3, 4));

// Dynamic code generation
var encodedCode = "114,101,116,117,114,110,32,52,50";
var codes = encodedCode.split(",");
var funcBody = "";
for (var i = 0; i < codes.length; i++) {
    funcBody += String.fromCharCode(parseInt(codes[i]));
}
print("Dynamic function body: " + funcBody);
var dynamicFunc = new Function(funcBody);
print("Result: " + dynamicFunc());

print("=== Function Constructor Tests Complete ===");

/*
**Expected Audit Output:**
```
[FUNC_CTOR] return a + b;
[FUNC_CTOR] return "Hello from dynamic function!";
[FUNC_CTOR] return x * y * z;
[DECODE] String.fromCharCode(72) -> H
[DECODE] String.fromCharCode(101) -> e
[DECODE] String.fromCharCode(116) -> t
... (more decodes)
[FUNC_CTOR] return 42
*/








