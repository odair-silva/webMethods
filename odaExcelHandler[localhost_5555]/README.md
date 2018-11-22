# [webMethods] ExcelReader to JSON

This is my implementation to read a xls file and to output a JSON with the contents.
There are two working services: one organizes the JSON as rows and the other organizes by the field in the first column of the file. 

Examples below:

---

#### Case 1



##### Input (xls):  

_ | A | B | C | ... | Last Column
-- | -- | -- | -- | -- | --
**1** | Field 1 | Field 2 | Field 3 | ... | Field n
**2** | a | b | c | ... | n
**3** | aa | bb | cc | ... | nn
... | ... | ... | ... | ... | ...
**Last Row** | aaa | bbb | ccc | ... | nnn




##### Output (JSON):

```json
 [{
 		"row 1": {
 			"Field 1": "a",
 			"Field 2": "b",
 			"Field 3": "c",
 			...,
 			"Field n": "n"
 		}
 	},

 	{
 		"row 2": {
 			"Field 1": "aa",
 			"Field 2": "bb",
 			"Field 3": "cc",
 			...,
 			"Field n": "nn"
 		}
 	},

 	...,

 	{
 		"row n": {
 			"Field 1": "aaa",
 			"Field 2": "bbb",
 			"Field 3": "ccc",
 			...,
 			"Field n": "nnn"
 		}
 	}
 ]
```
 
---

#### Case 2 



##### Input (xls):

_ | A | B | C | ... | Last Column
-- | -- | -- | -- | -- | --
**1** | | Field 1 | Field 2 | ... | Field n
**2** | Mary | b | c | ... | n
**3** | Evander | bb | cc | ... | nn
... | ... | ... | ... | ... | ...
**Last Row** | Henry | bbb | ccc | ... | nnn


Output (JSON): 

```json
 [{
		"Mary": {
			"Field 1": "b",
			"Field 2": "c",
			...,
			"Field n": "n"
		}
	},

	{
		"Evander": {
			"Field 1": "bb",
			"Field 2": "cc",
			...,
			"Field n": "nn"
		}
	},

	...,

	{
		"Henry": {
			"Field 1": "bbb",
			"Field 2": "ccc",
			...,
			"Field n": "nnn"
		}
	}
]
```
---

