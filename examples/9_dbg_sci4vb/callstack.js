function a(x){
  x++
  console.log("inside a(" + x + ")");
}

function b(x){
  x++
  console.log("inside b(" + x + ")");
  a(x)
}

function c(x){
  x++
  console.log("inside c(" + x + ")");
  b(x)
}

c(0)
























