# TestComplete_Auto-Part-Processing
This project takes parts in insights and goes through slicing a part (potentially parts) from a specific printer and materials.

How does it work?
Currently the user will 
  1) Open up insight
  2) Put the file the user wishes to slice into insight
  3) Run the main 'part_processing_script'
  4) Select which printer the user wishes to slice for
  5) Select the material/s that the user wishes to slice for
  6) Sit back and relax
  
TODO:
(in no particular order)
- being able to slice for only a specific slice height
  - will probably need a flag, and check for that flag at certain parts?
- being able to select multiple materials from a check box (instead of dropdown)
  - have a full list of materials and the ones not able to be selected are greyed out
