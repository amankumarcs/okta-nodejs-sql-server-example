var testSort= [
    { 'ATM ID': 'ARN9006', Deno: '1000', "three" : "three1" },
    { 'ATM ID': 'ARN9006', Deno: '500', "three" : "three2" },
    { 'ATM ID': 'ADU8001', Deno: '6000', "three" : "three2" },
    { 'ATM ID': 'AMN9048', Deno: '500', "three" : "three3" },
    { 'ATM ID': 'ADU8001', Deno: '7000', "three" : "three3" },
    { 'ATM ID': 'ADU8001', Deno: '10000', "three" : "three3" },
    { 'ATM ID': 'AMN9048', Deno: '200', "three" : "three3" },
    { 'ATM ID': 'AMN9048', Deno: '1000', "three" : "three3" },
    { 'ATM ID': 'ADU8001', Deno: '15000', "three" : "three3" },
    { 'ATM ID': 'ADU8001', Deno: 'Total', "three" : "three3" }
   
]


testSort.sort(function (a, b) {
    return a["ATM ID"].localeCompare(b["ATM ID"]) || a["Deno"] - b["Deno"];
});
let x = [];
let sortData = testSort.reduce((r, a) => {
    
    r[a["ATM ID"]] = [...r[a["ATM ID"]] || [], a];
    x.push(a["ATM ID"])
    return r;
   }, {});

   let data = []

var _ = require('underscore');
let uniQueIndex = _.uniq(x)
uniQueIndex.forEach(element => {
    console.log(sortData[element]);
    

   


});


//items.sort(function (a, b) {
  //return a.value - b.value;
//});
