# Patent "CitedBy" Retriever
A tool for retrieving patent "Cited By" information from Google Patent

![Retriever](https://github.com/Zack-Cheng/Patent-CitedBy-Retriever/blob/master/PatentCitedByRetriever.PNG)

Input example (Excel `.xlsx`):

**Section1**|**Section2**|**Section3**
:-----:|:-----:|:-----:
7973130|9478942|3108359
2478908|3234540|3098160
3413480|9862947|3257434
3009671|2872622|3257433

Each column in the first row of the input is user-defined collection name.

The second and the following rows are patent publication numbers without country code (the first two letters "US") and the kind codes (A1/B2 etc suffix) in each collection.

This tool will collect "Cited By" information for each patent and add up the number of patents that cite it.
![Cited by](https://github.com/Zack-Cheng/Patent-CitedBy-Retriever/blob/master/CitedBy.png)
