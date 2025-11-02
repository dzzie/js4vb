// COMPREHENSIVE AUDIT MODE TEST SCRIPT
// This script triggers all enumAuditEvents in CInterpreter


// ========================================
// 1. aeEval - eval() call
// ========================================
eval("var x = 42;");
eval("1 + 1");

// ========================================
// 2. aeDecode - Decode functions
// ========================================
unescape("%48%65%6C%6C%6F");
decodeURI("Hello%20World");
decodeURIComponent("Hello%20World%21");

// Test String.fromCharCode
var hello = String.fromCharCode(72, 101, 108, 108, 111);  // "Hello"
var cmd = String.fromCharCode(99, 109, 100);
//print(hello + ' ' + cmd)

// ========================================
// 3. aeBracketAccess - Computed property access
// ========================================
var obj = { foo: "bar", test: 123 };
var propName = "foo";
var result = obj[propName];
var arr = [1, 2, 3];
var idx = 1;
var val = arr[idx];

// Dynamic property access
var shell = "WScript";
var method = "Shell";
// This will trigger bracket access even if object doesn't exist
try {
    var x = obj["dynamicProp"];
} catch(e) {}

// ========================================
// 4. aeCOMCall - COM method calls
// ========================================
var wsh = new ActiveXObject("WScript.Shell");
wsh.Run("cmd.exe /c echo test");
wsh.Exec("notepad.exe");

var fso = new ActiveXObject("Scripting.FileSystemObject");
fso.FileExists("C:\\test.txt");

// ========================================
// 5. aeActiveX - ActiveXObject creation
// ========================================
var shell2 = new ActiveXObject("WScript.Shell");
var network = new ActiveXObject("WScript.Network");
var fso2 = new ActiveXObject("Scripting.FileSystemObject");
var stream = new ActiveXObject("ADODB.Stream");

// ========================================
// 8. aeFunctionConstructor - Function constructor
// ========================================
// NOTE: This would need to be implemented in your interpreter
// var dynamicFunc = new Function("a", "b", "return a + b;");
// var result = dynamicFunc(1, 2);

// 1. Decode hex string using parseInt
var codes = "48656c6c6f";
var decoded = "";
alert(codes.length)
for (var i = 0; i < codes.length; i += 2) {
    var hexByte = codes.substr(i, 2);
    var charCode = parseInt(hexByte, 16);  // aeParse
    c = String.fromCharCode(charCode);
    //decoded += c
}
print('decoded: '+decoded)

// 2. XOR decode
var xorEncoded = [0x7a, 0x77, 0x7e, 0x7e, 0x73];
var xorKey = 0x12;
var xorDecoded = "";
for (var j = 0; j < xorEncoded.length; j++) {
    var decodedByte = xorEncoded[j] ^ xorKey;  // aeXOR
    xorDecoded += String.fromCharCode(decodedByte);
}

// 3. Unescape
var escaped = unescape("%57%53%63%72%69%70%74%2E%53%68%65%6C%6C");  // aeDecode

// 4. Dynamic property access
var progId = "WScript";
var component = "Shell";
var fullProgId = progId + "." + component;

// 5. Create COM object dynamically
var shell3 = new ActiveXObject(fullProgId);  // aeActiveX

// 6. Build command using bracket notation
var commands = { cmd: "cmd.exe", note: "notepad.exe" };
var cmdKey = "cmd";
var command = commands[cmdKey];  // aeBracketAccess

// 7. Execute via COM
shell3.Run(command);  // aeCOMCall

// 8. Eval decoded code
var encodedCode = "76617220793d3432";  // hex for "var y=42"
var decodedCode = "";
for (var k = 0; k < encodedCode.length; k += 2) {
    decodedCode += String.fromCharCode(parseInt(encodedCode.substr(k, 2), 16));
}
eval(decodedCode);  // aeEval

print("All audit events triggered!");
 











