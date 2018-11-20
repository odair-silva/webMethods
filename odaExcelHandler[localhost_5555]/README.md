ExcelHandler

This is my implementation to read from a xls file through webMethods and having a JSON as output.
There are two working services: one organizes the JSON as rows and the other organizes by the field in the first column of the file.

1st case:  
| x1 | x2 | x3 | ... | xn |  
| o | o | o | ... | o |  
| o | o | o | ... | o |  
  .   .   .         .  
  .   .   .         .  
  .   .   .         .  
| o | o | o | ... | o |  

output: \[{"row 1" :{"x1" : "o", "x2" : "o", "x3" : "o" ...}}, {"row 2" :{"x1" : "o", "x2" : "o", "x3" : "o" ...}},...\]  

2nd case:  
| - | x1 | x2 | ... | xn |  
| y1 | o | o | ... | o |  
| y2 | o | o | ... | o |  
  .   .   .         .  
  .   .   .         .  
  .   .   .         .  
| yn | o | o | ... | o |  

output: \[{"y1" :{"x1" : "o", "x2" : "o", "x3" : "o" ...}}, {"y2" :{"x1" : "o", "x2" : "o", "x3" : "o" ...}},...\]  
