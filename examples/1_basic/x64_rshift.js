// Simple Int64 >>> test
print("=== Int64 >>> Test ===");

var bigNum = 0x8000000000000000;  // Requires Int64
print("Big number: " + hex(bigNum));

var shifted = bigNum >>> 4;
print("After >>> 4: " + hex(shifted));

var bigNum2 = 0xFFFFFFFF00000000;
bigNum2 >>>= 8;
print("After >>>= 8: " + hex(bigNum2));

print("=== Test Complete ===");

/*
=== Int64 >>> Test ===
Big number: 0x8000000000000000
After >>> 4: 0x0800000000000000
After >>>= 8: 0x00FFFFFFFF000000
=== Test Complete ===
*/

