//date handling contributed by yokesee thanks!

var now = new Date();
print("Current date: " + now.toString());
print("Current date: " + now);  
print("Date.now() static: " + Date.now()); 

print("Year: " + now.getFullYear());
print("Month: " + now.getMonth());
print("Day: " + now.getDate());
print("Timestamp: " + now.getTime());

var User= {
    name: "Juan",
    DateRegistred: new Date().toString()
};

print(JSON.stringify(User, null, 2));

var dateStr = new Date().toString();
print("Date in string: " + dateStr);