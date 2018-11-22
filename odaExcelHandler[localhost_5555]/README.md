[webMethods] ExcelReader to JSON

This is my implementation to read a xls file and to output a JSON with the contents.
There are two working services: one organizes the JSON as rows and the other organizes by the field in the first column of the file. Examples below:

---

1st case



Input (xls):  

| x1 | x2 | x3 | ... | xn |  
|  o  |  o  |  o  | ... |  o  |  
|  o  |  o  |  o  | ... |  o  |  
​     .       .       .               .  
​     .       .       .               .  
​     .       .       .               .  
|  o  |  o  |  o  | ... |  o  |  



Output (JSON):

 \[{"row 1" :{"x1" : "o", "x2" : "o", "x3" : "o" , ..., "xn" : "o"}}, 

   {"row 2" :{"x1" : "o", "x2" : "o", "x3" : "o" , ..., "xn" : "o"}},

   ...,

   {"row n" :{"x1" : "o", "x2" : "o", "x3" : "o" , ..., "xn" : "o"}}\]  

---

2nd case:  



Input (xls):

|      | x1 | x2 | ... | xn |  
| y1 |  o  |  o  | ... |  o  |  
| y2 |  o  |  o  | ... |  o  |  
​     .       .       .               .  
​     .       .       .               .  
​     .       .       .               .  
| yn |  o  |  o  | ... |  o  |

Output (JSON): 

 [{"y1" :{"x1" : "o", "x2" : "o", ..., "xn" : "o"}}, 

  {"y2" :{"x1" : "o", "x2" : "o", ..., "xn" : "o"}},

   ...,

  {"yn" :{"x1" : "o", "x2" : "o", ..., "xn" : "o"}}]  

---

