﻿"""
Written by: Andy Groth
Modified date: 2/1/2023

Script for processing parts. 
Second time writing since first one 
was accidentally deleted.

"""
import insight
import forms

click = insight.click
log = Log.Message
key_stroke = Sys.Process("insight").Window("TkTopLevel", "*", 1).Keys

# These variables will be saved as the User selects from the dropdown menu 
#   on the form
USER_SYSTEM = ""
USER_MATERIALS = ""
PART_NAME = ""

def wait():
  return insight.wait(insight.ONE_SECOND)


def main():
  """
  This script will press buttons and save files based on which type of printer the user will select
  from the 'part processing form' once that is made. The part will need to be pulled into 
  Insight and sized/oriented properly and then the script will go through and process them.
  """
  # Have the user select which gender of printer and material they are wanting to  
  #   slice from the form
  insight.insight_to_front()
  USER_SYSTEM, USER_MATERIALS = forms.open_processing_userform()
  
  # Getting the name of the part we are slicing
  part_from_insight = Sys.Process("insight").Window("TkTopLevel", "*", 1).WndCaption
  PART_NAME = part_from_insight[10:]
  
  
  ### Starting ###
  # A message to the User so that they understand what must be done before the scrip starts
  BuiltIn.ShowMessage("Before clicking 'OK' make sure that you have a part populated in Insight " \
  "and that it's sized and rotated correctly. Once the part is good, press 'OK' to continue.")
  
  # Setting up insight to the 'default' position
  insight.pre_processing()
  
  
  ### Printer Selection ###
  # Navigating to the 'Configure Modeler'
  click(insight.configure_modeler())
  insight.Looping_Functions.configure_modeler_checking() 
  wait()
  
  # Setting the printer type
  insight.Looping_Functions.set_printer(USER_SYSTEM)
  
  
  ### Material Selection ###
  # Checking to see what material the user selected for 
  if USER_MATERIALS == "ALL":
    # The list of materials/tips is located in the 'Full-List' excel doc in the same folder
    #   as this project
    
    # Getting the full list of materials 
    all_materials = forms.get_materials_from_excel(USER_SYSTEM)
    
    ### Slicing Process ###
    # While loop to cycle through and select each material
    for current_material in all_materials:
      #Log.Message(f"current_material is '{current_material}'")
      tab_number = insight.Looping_Functions.loop_slicing(USER_SYSTEM, current_material, PART_NAME)
      
    
  # We want to slice for a specific material 
  else:
    ### Slicing Process ###  
    # This function takes the printer selected and material selected and loops through the process
    #   of selecting each combination, slicing the part, and waiting for the part to finish
    #   slicing and then repeating the process
    tab_number = insight.Looping_Functions.loop_slicing(USER_SYSTEM, USER_MATERIALS,PART_NAME)
