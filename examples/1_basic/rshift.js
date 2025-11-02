// Test unsigned right shift edge cases
print("=== Unsigned Right Shift Tests ===");

// Test 1: Positive number (no difference from >>)
var a = 16;
print("16 >>> 2 = " + (a >>> 2));  // Should be 4

// Test 2: Negative number (KEY TEST!)
var b = -1;
print("-1 >>> 0 = " + (b >>> 0));  // Should be 4294967295 (2^32 - 1)

// Test 3: Large negative
var c = -2147483648;  // MIN_INT32
print("-2147483648 >>> 0 = " + (c >>> 0));  // Should be 2147483648

// Test 4: With actual shift
var d = -1;
print("-1 >>> 1 = " + (d >>> 1));  // Should be 2147483647 (2^31 - 1)

// Test 5: Shift amount > 31 (should mask to 5 bits)
var e = 16;
print("16 >>> 34 = " + (e >>> 34));  // Same as 16 >>> 2 = 4

// Test 6: Compound assignment
var f = -1;
f >>>= 1;
print("-1 >>>= 1 result: " + f);  // Should be 2147483647

print("=== Tests Complete ===");

