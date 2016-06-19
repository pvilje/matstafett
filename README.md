# matstafett
Read a list of participants and generate a lineup for "matstafett"
(No idea what matstafett is called in english)

**Basic concept**
A certain number of people/couples sign up. The participant total MUST be a factor of 3, and at least 9.
Each person/couple is assigned a part of a dinner to host (starter, main course or desert).
Each dinner will consist of a host and two guests. The meal will take place at the host's home.

The dinners start at the starter hosts, each host will have 2 guests. 
They then split up and move on to a main course host where there also will be two guests, same thing for the desert. 

The goal of it all is:
* For everyone to get a part of the dinner to prepare.
* Everyone will attend to a starter, main course and desert.
* Everyone will meet new people in each meal. 
* (3 meals with 3 participants at each meal, regardless of the amount of participants each participant will have met 8 other participants at the end, each participant will have hosted a meal, each participant will have visited two others home)

**The program:**
v.1 will probably be in swedish, later version will include translation files... 
Select an Excel (.xlsx) or notepad (.txt) document which only should contain a list of participants, nothing else.
The script will calculate a setup where all participants will:
* Get a certain part of the dinner to prepare.
* Meet new people at every part of the dinner
* Save the result in the same file type as was submitted to the script.
* If an excel file is submitted, save the result in a new sheet (same file) nicely formatted.
* If a txt file is submitted, save the result in a new txt document.
