"""
Written by: Andy Groth
Modified date: 2/1/2023

Script for processing parts. 
Second time writing since first one 
was accidentally deleted.

"""
import insight
import forms
import os
import re

click = insight.click
log = Log.Message

# These variables will be saved as the User selects from the dropdown menu 
#   on the form
USER_SYSTEM = ""
USER_MATERIALS = ""
PART_NAME = ""
DIRECTORY = ""

def main():
  """
  This script will press buttons and save files based on which type of printer the user will select
  from the 'part processing form' once that is made. The part will need to be pulled into 
  Insight and sized/oriented properly and then the script will go through and process them.
  """
  # Have the user select which gender of printer and material they are wanting to  
  #   slice from the form
  insight.insight_to_front()
  USER_SYSTEM, USER_MATERIALS, DIRECTORY = forms.open_processing_userform()
  #Log.Message(f"DIRECTORY is {DIRECTORY}")
  folder_path = DIRECTORY
    
  # A message to the User so that they understand what must be done before the scrip starts
  BuiltIn.ShowMessage("Before clicking 'OK' make sure that you have a part placed in " \
    "the '" + folder_path + "' folder. Once this is good, press 'OK' to continue.")
  insight.insight_to_front() 
   
  # Getting how many stls and names of stls from the files 
  number_STLs, stl_files = insight.Looping_Functions.get_stl_files_info(folder_path)
  
  # Starting the loop for going through each stl
  for x in range(number_STLs):
    insight.Looping_Functions.open_up_stl(stl_files[x], folder_path)
    
    # Wait for file opening finish
    while Sys.Process("insight").WaitChild('Window("TkTopLevel", "Loading STL file", 1)', 1000).Exists:
      insight.wait()
      # if something pop-ups, click yes
      if Sys.Process("insight").WaitChild('Window("#32770", "Insight", 1)', 1000).Exists:
        click(insight.overwrite.overwrite_box_yes())  
    
        
    ### Starting ###
    Log.Message("Starting")
    # Getting the name of the part we are slicing
    part_from_insight = Sys.Process("insight").Window("TkTopLevel", "*", 1).WndCaption
    PART_NAME = part_from_insight[10:]
    
    # Setting up insight to the 'default' position
    insight.pre_processing()
  

    ### Printer Selection ###
    Log.Message("Printer Selection")
    # Navigating to the 'Configure Modeler'
    insight.open_configure_modeler()
    #insight.Looping_Functions.configure_modeler_checking() 
    insight.wait()
  
    # Setting the printer type
    insight.Looping_Functions.set_printer(USER_SYSTEM)
  
  
    ### Material Selection ###
    Log.Message("Material Selection")
    # Checking to see what material the user selected for 
    if USER_MATERIALS == "ALL":
      # The list of materials/tips is located in the 'Full-List' excel doc in the same folder
      #   as this project
    
      # Getting the full list of materials 
      all_materials = forms.get_materials_from_excel(USER_SYSTEM)
    
      ### Slicing Process ###
      # While loop to cycle through and select each material
      for current_material in all_materials:
        tab_number = insight.Looping_Functions.loop_slicing(USER_SYSTEM, current_material, PART_NAME)
      
    
    # We want to slice for a specific material 
    else:
      ### Slicing Process ###  
      # This function takes the printer selected and material selected and loops through the process
      #   of selecting each combination, slicing the part, and waiting for the part to finish
      #   slicing and then repeating the process
      tab_number = insight.Looping_Functions.loop_slicing(USER_SYSTEM, USER_MATERIALS,PART_NAME)

      

