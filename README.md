# TestComplete_Auto-Part-Processing
This project takes parts in insights and goes through slicing stl's for a specific printer and material/s.
A copy of the project is located on the sharedrive here: S:\Dept\RandD\Test Engineering (Software)\TestComplete_Auto-Part-Processing


How does it work?
- This script will slice for each slice height and support combination for the selected materials. If all of the materials are selected it will do this for all of them. It will go through all of the materials for one stl, and then move on to the next until they are all finished.
- Since we do not have access to certain information through insight, the Excel document 'Full-List' is what the script uses to "see" information. This information is divided into a list of printers that are currently supported, a sheet for each printer, and the materials/slice height/tip sizes for each printer. This data is arranged in the same way as it appears on insight. 
- **NOTE: IF THIS DATA MOVES POSITIONS IN INSIGHT (ie. a new printer is added in the list), IT WILL BREAK THE SCRIPT SO WE SHOULD BE OCCASIONALLY CHECKING THE DOCUMENTS VALIDITY

Using the Script. 
Once the User has the project copied they will;
  1) Put the parts they wish to slice all into the same folder.
  2) Open up insight
  3) Run the 'part_processing_script' in TestComplete or TestExcicute 
  4) Select where the stl's are located (this will default to the 'Parts' folder on the sharedrive here 'S:\Dept\RandD\Test Engineering (Software)\TestComplete_Auto-Part-Processing\Parts')
  5) Select which printer the user wishes to slice for
  6) Select the material/s that the user wishes to slice for
  7) Sit back and relax

This script will slice for each slice height and support combination for the selected materials. If all of the materials are selected it will do this for all of them. It will go through all of the materials for one stl, and then move on to the next until they are all finished.


  
TODO:
(in no particular order)
- being able to slice for only a specific slice height
  - will probably need a flag, and check for that flag at certain parts?
- being able to select multiple materials from a check box (instead of dropdown)
  - have a full list of materials and the ones not able to be selected are greyed out
