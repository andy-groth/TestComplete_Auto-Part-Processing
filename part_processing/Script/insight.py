"""
This is the Insight library that is holding the mappings for buttons in insight as well as
helpful functions

"""
import forms
log = Log.Message

########      Global Variables        #########
# Times
ONE_SECOND = 1000
FIVE_SECONDS = 5 * ONE_SECOND
TEN_SECONDS = 10 * ONE_SECOND
ONE_MINUTE = 60 * ONE_SECOND
FIVE_MINUTES = 5 * ONE_MINUTE
TEN_MINUTES = 10 * ONE_MINUTE

# this variable is one that represents the max number of printers
PRINTER_RANGE = 40

# Project name of the repo
PROJECT_NAME = "Part Processing"
PROJECT_PATH = "C:\Automation\Part_Processing"
SHAREDRIVE_PATH = "S:\Dept\RandD\Test Engineering (Software)\TestComplete - Auto Part Processing\Current"


########      START OF FUNCTIONS       #########  

# a mask for TestCompletes Activate function for easier 
def insight_to_front():
  """
  This function is a mask to make understanding the code easier
  """
  return Sys.Process("insight").Window("TkTopLevel", "*", 1).Activate()

  
# A mask for TestComplete's delay function.
def wait(time: int):
	aqUtils.Delay(time)

	
# The process that happens to get insight to a normalized state
def pre_processing():
  """
  The goal of this pre_processing is to get insight to a 'normalized'
  position so that we have the same starting point every time the script
  is being ran.
  """
  
  # focus insight for start of the script
  insight_to_front()
  wait(ONE_SECOND)
  
  # navigating into the configure modeler 
  click(insight.configure_modeler())
  wait(ONE_SECOND)

  
  #insight_to_front()
  Looping_Functions.configure_modeler_checking() 
  
  # setting modeler type availability
  click(config_mod_screen.modeler_type_availability())
  insight_to_front()
  click(config_mod_screen.modeler_type())
  wait(ONE_SECOND)
  
  # this next chunk is added to 'clear' the model/support material in case the printer is 
  # configured to 900mc
  config_mod_screen.modeler_type_button().Keys("[Enter]")
  config_mod_screen.modeler_dropdown().Keys("[Down]")
  config_mod_screen.modeler_dropdown().Keys("[Enter]")
  wait(ONE_SECOND)
  config_mod_screen.modeler_type_button().Keys("[Enter]")
  
  # this loop brings the printer gender back to the top (40 is an arbitrary number that will probably 
  # not be exceeded in the close future and doesn't add much time to the test)
  config_mod_screen.modeler_dropdown().Keys("[Up]"*PRINTER_RANGE + "[Enter]")
  wait(ONE_SECOND)
  
  # tabbing down from the system type to the check button and clicking 
  # the parent button (the child button was "disapearing" causing issues
  config_mod_screen.modeler_type().Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]")
  Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
  wait(ONE_SECOND)
  
  
  
  
#####    Testing Helper Functions
def assertion(condition: bool, condition_message: str):
	"""
	Function to assert a general condition and logs it as an assertion step.
	Args:
		condition (bool): The condition to verify.
		condition_message (str): The condition that is being verified.

	Returns:
	Nothing
	"""

	if condition:
		Log.Checkpoint("Assertion Success: " + condition_message)
	else:
		Log.Error("Assertion Failure: " + condition_message)
	
			
		
def yes_or_no_dialog(yes_or_no_expected: bool, question: str) -> bool:
	"""
	Generic function to ask the tester a yes or no question, returns the response.
	Args:
		yes_or_no_expected (bool): Condition value expected from the tester.
		question (str): Question to ask the tester.

	Returns:
	bool. Response from the tester.
	"""
	# Expected is the integer value of "Yes" if true, else it's the No value.
	expected = dialog_button_values["Yes"] if yes_or_no_expected else dialog_button_values["No"]
	return expected == BuiltIn.MessageDlg(question, mtConfirmation, MkSet(mbYes, mbNo), 0)
	
			
	
def click(ui_element):
	"""
	Generic function to click a UI element with a one second delay.
	Args:
	ui_element (TestComplete Element): The element to click.

	Returns:
	Nothing
	"""

	wait_ui_timeout = FIVE_MINUTES
	is_visible = ui_element.WaitProperty("VisibleOnScreen", "True", wait_ui_timeout)

	if not is_visible:
		question = "The automation is attempting to click a UI element." + \
				   "However, it is not enabled or not visible on screen." \
				   "Continue waiting and try again? "
		if yes_or_no_dialog(True, question):
			ui_element.click()

		else:
			error = "Cannot click(): Element not enabled or visible."
			Log.error(error)
	ui_element.click()		
		
	
	

  	
	
  
  
### LOOPING FUNCTIONS 
class Looping_Functions():
  """ 
  This is the class for getting the movement information for each material type
  tip, slice etc. This will also be for naming the parts and resetting insight 
  positions
  """
  
  def reset_pos(what_to_reset:str):
    """
    This function takes which dropdown to reset as a string and resets that to
    the top position.
    """
    x = 0
    dropdown = what_to_reset.lower()
    
    if dropdown == "model": 
      click(config_mod_screen.model_material())
      
      config_mod_screen.model_material_dropdown().Keys("[Up]"*PRINTER_RANGE + "[Enter]")
      log("Reset Model Material")
      
    elif dropdown ==  "support":
      click(config_mod_screen.support_material())
      config_mod_screen.support_material_dropdown().Keys("[Up]"*PRINTER_RANGE + "[Enter]")
      log("Reset Support Material")
        
    elif dropdown == "slice":
      click(config_mod_screen.slice_height())
      config_mod_screen.slice_height_dropdown().Keys("[Up]"*PRINTER_RANGE + "[Enter]")
      log("Reset Slice Height")
        
    elif dropdown == "model_tip":
      click(config_mod_screen.model_tip())
      config_mod_screen.model_tip().Keys("[Enter]")
      config_mod_screen.model_tip_dropdown().Keys("[Up]"*PRINTER_RANGE + "[Enter]")
      log("Reset Model Tip")
        
    elif dropdown == "support_tip":
      click(config_mod_screen.support_tip())
      config_mod_screen.support_tip_dropdown().Keys("[Up]"*PRINTER_RANGE + "[Enter]")
      log("Reset Support Tip")
        
    else:
      BuiltIn.ShowMessage(f"Reset automation failed: wrong function input. Please select the" 
      "correct dropdown, and click on the first item at the stop.")
    
  
      
  def loop_slicing(printer_type: str, model_material: str, part_name: str):
    """
    This function is one that will take in printer type and model material and
    do the selection of model, support, slice height, and model tip. slice the file 
    and then wait for the prat to be processed before looping again.
    """
    # Save file name variables
    savefile_part_name = part_name
    save_file_printer = printer_type
    savefile_model_name   = model_material
    savefile_support_name = ""
    savefile_slice_height = ""
    savefile_model_tip    = ""
    savefile_support_tip  = ""
    
    # Strings from the excel doc
    model_list_str        = ""
    support_list_str      = ""
    slice_list_str        = ""
    model_tip_list_str    = ""
    support_tip_list_str  = ""
    tab_list_str          = ""
    
    model_list          = []
    support_list        = []
    slice_list          = []
    model_tip_list      = []
    support_tip_list    = []
    tab_list            = []
    
    model_command       = None
    support_command     = None
    slice_command       = None
    model_tip_command   = None
    support_tip_command = None
    tab_number          = None
    
    
    ### Starting Slicing Process ###
    # Open modeler window
    click(insight.configure_modeler())
    Looping_Functions.configure_modeler_checking() 
    
    click(config_mod_screen.modeler_type_availability())
    click(config_mod_screen.modeler_type())
   
    
    # Getting model command to select the specific material
    model_command = Looping_Functions.get_model_move(printer_type, model_material)
    #Log.Message(f"model_command is '{model_command}'")
    
    ### Select Model Material ###
    # Find model material from dropdown 
    insight_to_front()
      
    # Reset model material
    Looping_Functions.reset_pos("model")
    
    # Select model material
    click(config_mod_screen.model_material())
    config_mod_screen.model_material_dropdown().Keys("[Down]"*model_command)
    config_mod_screen.model_material_dropdown().Keys("[Enter]")
    log("Model Material Selected")
      
      
    # Getting info strings for the selected material
    support_list_str, slice_list_str, model_tip_list_str, support_tip_list_str, tab_list_str = forms.Navigation.get_material_info(printer_type, model_material)
    
    # Log.Message(f"support_list_str is {support_list_str}")
    # Log.Message(f"slice_list_str is {slice_list_str}")
    # Log.Message(f"model_tip_list_str is {model_tip_list_str}")
    # Log.Message(f"support_tip_list_str is {support_tip_list_str}")
    # Log.Message(f"tab_list_str is {tab_list_str}")
    
    # Turning the strings into lists to go through
    support_list     = forms.string_to_list(support_list_str)
    #log("support list to string.")
    slice_list       = forms.string_to_list(slice_list_str)
    #log("slice list to string.")
    model_tip_list   = forms.string_to_list(model_tip_list_str)
    #log("model tip list to string.")
    support_tip_list = forms.string_to_list(support_tip_list_str)
    #log("support tip list to string.")
    tab_list         = forms.string_to_list(tab_list_str)  
    #log("tab list to string.")
    
    ## counter for testing/debugging
    total_number_sliced = 1
    
    ### Select Support Material ###
    # set support command
    support_command = 0
    
    for support_material in support_list:
      #Log.Message(f"##### Select Support Material #####")
      #Log.Message(f"support_material is '{support_material}'") 
        
      # this is to keep the variable known for future assigning of variables
      support_command_used = support_command
      
      #reset support material
      Looping_Functions.reset_pos("support")
      
      #select support material
      click(config_mod_screen.support_material())
      config_mod_screen.support_material_dropdown().Keys("[Down]"*support_command)
      config_mod_screen.support_material_dropdown().Keys("[Enter]")
        
      log("Support Material Selected")  
      # set support name for save file
      savefile_support_name = support_list[support_command][0]
      #debug
      #Log.Message(f"savefile_support_name is '{savefile_support_name}'")
      
      # incrament the command incase there are more
      support_command += 1
      
      #Log.Message(f"support_command is {support_command}")
      #Log.Message(f"support_command_used is {support_command_used}")
      
      
      ### Slice Height ###
      slice_command = 0
      
      # Set slice height command
      for slice_item in slice_list[support_command_used]:
        #Log.Message(f"##### Slice Tip Loop #####")
        #Log.Message(f"slice_item is '{slice_item}'")
          
          
        # This is to keep the variable known for future assigning of variables
        slice_command_used = slice_command
        
        # Reset slice height
        Looping_Functions.reset_pos("slice")
        
        # Select slice height
        click(config_mod_screen.slice_height())
        config_mod_screen.slice_height_dropdown().Keys("[Down]"*slice_command)
        config_mod_screen.slice_height_dropdown().Keys("[Enter]")
        log("Slice Height Selected")
        
        # Set slice height for save file
        savefile_slice_height = slice_list[support_command_used][slice_command]
        #debug
        #Log.Message(f"savefile_slice_height is '{savefile_slice_height}'")
        
        #incrament the slice command
        slice_command += 1
        
        
        
        ### Model Tip ###
        # set support command
        model_tip_command = 0
        support_tip_command = 0
        #Log.Message(f"model_tip_list is {model_tip_list}")
        #Log.Message(f"model_tip_list[{slice_command_used}] is {model_tip_list[support_command_used]}")
        #Log.Message(f"model_tip_list[{support_command_used}][{slice_command_used}] is {model_tip_list[support_command_used][slice_command_used]}")
        
        # this is checking to see if the item within the model_tip_list is a nested list
        #   aka needs to be looped through
        if isinstance(model_tip_list[support_command_used][slice_command_used], list) is True:

          for model_tip_item in model_tip_list[support_command_used][model_tip_command]:
            #Log.Message(f"##### Model Tip Loop #####")
            #Log.Message(f"model_tip_item is {model_tip_item}")

            # this is to keep the variable known for future assigning of variables
            model_tip_command_used = model_tip_command
            #Log.Message(f"model_tip_command is {model_tip_command}")
            #Log.Message(f"model_tip_command_used is {model_tip_command_used}")

            # reset model_tip
            Looping_Functions.reset_pos("model_tip")
            
            # select model_tip
            click(config_mod_screen.model_tip())
            config_mod_screen.model_tip().Keys("[Enter]")
            config_mod_screen.model_tip_dropdown().Keys("[Down]"*model_tip_command)
            config_mod_screen.model_tip_dropdown().Keys("[Enter]")
            
            log("Model Tip Selected")
            
            # set slice height for save file
            savefile_model_tip = model_tip_item
            #debug
            #Log.Message(f"model_tip_list is {model_tip_list[model_tip_command_used]}")
            #Log.Message(f"savefile_model_tip is '{savefile_model_tip}'")
            
            #incrament the model tip command
            model_tip_command += 1

            
            ### Support Tip ###
            #Log.Message(f"support_tip_list[{support_command_used}][{slice_command_used}] is {support_tip_list[support_command_used][slice_command_used]}")
            if isinstance(support_tip_list[support_command_used][slice_command_used], list) is True:

              for support_tip_item in support_tip_list[support_command_used][support_tip_command]:
                #Log.Message(f"+++++ MODEL TIP LOOP AND SUPPORT TIP LOOP HIT +++++")
                #Log.Message(f" support_tip_item is {support_tip_item}")
                #Log.Message(f"support_tip_list[{support_command_used}][{slice_command_used}] is {support_tip_list[support_command_used][slice_command_used]}")
                
                # this is to keep the variable known for future assigning of variables
                support_tip_command_used = support_tip_command

                #debug
                #Log.Message(f"support_tip_item[0] is {support_tip_item[0]}")
            
            
                # reset support_tip
                Looping_Functions.reset_pos("support_tip")
            
                # select support_tip
                click(config_mod_screen.support_tip())
                config_mod_screen.support_tip_dropdown().Keys("[Down]"*support_tip_command)
                config_mod_screen.support_tip_dropdown().Keys("[Enter]")
              
                log("Support Tip Selected")
                # set support tip name for save file
                savefile_support_tip = support_tip_item

                #debug
                #Log.Message(f"+++++ savefile_support_tip is '{savefile_support_tip}'")
                #incrament the support tip command
                support_tip_command += 1
              
              
            # there is a choice for MATERIAL_TIP but no choice for SUPPORT_TIP
            else:
              #Log.Message(f"+++++ MODEL TIP LOOP HIT AND no support tip loop -----")
              savefile_support_tip = support_tip_list[support_command_used][slice_command_used]
              #debug
              #Log.Message(f"savefile_support_tip is '{savefile_support_tip}'")


            ### Tab Commands ###
            #Log.Message(f"model loop tab +++++")
            tab_number = tab_list[support_command_used][slice_command_used]
            tab_number = int(tab_number)
            #Log.Message(f"tab_number is '{tab_number}'")
            
            config_mod_screen.modeler_type().Keys("[Tab]" * tab_number)
            #log(f"tab down")
            Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
            #log(f"press ok")

            if Sys.Process("insight").WaitChild('Window("#32770", "Warning", 1)', 3000).Exists:
              click(insight.warning_window_yes())


            ### Slice Part ###
            total_number_sliced = Looping_Functions.slice_part_helper(total_number_sliced, save_file_printer, savefile_model_name, savefile_support_name, savefile_model_tip, savefile_support_tip, savefile_slice_height, part_name)
            #Log.Message(f"Slicing Part Started...")
            
            # open modeler window
            click(insight.configure_modeler())
            Looping_Functions.configure_modeler_checking() 
            #Log.Message(f"open config mod")
        
            
        # there is nothing for the model_tip to select
        else:
          # set model tip for save file
          savefile_model_tip = model_tip_list[support_command_used][slice_command_used]
          #debug
          #Log.Message(f"----- no model tip loop hit -----")
          #Log.Message(f"model_tip_list[{support_command_used}][{slice_command_used}] is model_tip_list[{support_command_used}][{slice_command_used}]")
          #Log.Message(f"savefile_model_tip is '{savefile_model_tip}'")

        
          ### Support Tip else ###
          support_tip_command = 0
          #Log.Message(f"support_tip_list[{support_command_used}][{slice_command_used}] is {support_tip_list[support_command_used][slice_command_used]}")
          if isinstance(support_tip_list[support_command_used][slice_command_used], list) is True:

            for support_tip_item in support_tip_list[support_command_used][support_tip_command]:
              #Log.Message(f"+++++ MODEL TIP LOOP AND SUPPORT TIP LOOP HIT +++++")
              #Log.Message(f" support_tip_item is {support_tip_item}")
              #Log.Message(f"support_tip_list[{support_command_used}][{slice_command_used}] is {support_tip_list[support_command_used][slice_command_used]}")
              
              # this is to keep the variable known for future assigning of variables
              support_tip_command_used = support_tip_command

              #debug
              #Log.Message(f"support_tip_item[0] is {support_tip_item[0]}")
          
          
              # reset support_tip
              Looping_Functions.reset_pos("support_tip")
          
              # select support_tip
              click(config_mod_screen.support_tip())
              config_mod_screen.support_tip_dropdown().Keys("[Down]"*support_tip_command)
              config_mod_screen.support_tip_dropdown().Keys("[Enter]")
            
              log("Support Tip Selected")
              # set support tip name for save file
              savefile_support_tip = support_tip_item
              #debug
              #Log.Message(f"+++++ savefile_support_tip is '{savefile_support_tip}'")
              #incrament the support tip command
              support_tip_command += 1
            
              
          #there is no choice for support tip
          else:
            #Log.Message(f"----- no model tip loop hit and no support loop -----")
            savefile_support_tip = support_tip_list[support_command_used][slice_command_used]
            #debug
            #Log.Message(f"support_tip_list[support_command_used][slice_command_used] is support_tip_list[{support_command_used}][{slice_command_used}]")
            #Log.Message(f"savefile_support_tip is '{savefile_support_tip}'")
              
            
          ### Tab Commands else ###
          #Log.Message(f"Tab no material------")
          tab_number = tab_list[support_command_used][slice_command_used]
          tab_number = int(tab_number[0])
          #Log.Message(f"tab_number is '{tab_number}'")
          
          config_mod_screen.modeler_type().Keys("[Tab]" * tab_number)
          #log(f"tab down")
          Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
          #log(f"press ok")
          
          # Changing the config will delete existing slices
          if Sys.Process("insight").WaitChild('Window("#32770", "Warning", 1)', 3000).Exists:
            click(insight.warning_window_yes())
            
            
          ### Slice Part ###
          #logging and slicing helper function
          total_number_sliced = Looping_Functions.slice_part_helper(total_number_sliced, save_file_printer, savefile_model_name, savefile_support_name, savefile_model_tip, savefile_support_tip, savefile_slice_height, part_name)
          
          
          # open modeler window
          click(insight.configure_modeler())
          Looping_Functions.configure_modeler_checking() 
          
          
          
    log("end of function")
     
    # Added this to the end of each pass through so we know exactly where we are 
    click(config_mod_screen.slice_height())
    config_mod_screen.slice_height_dropdown().Keys("[Down]"*slice_command)
    config_mod_screen.slice_height_dropdown().Keys("[Enter]")
    config_mod_screen.modeler_type().Keys("[Tab]" * tab_number)
    
    # closing the configure modeler    
    Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
    wait(ONE_SECOND)
    # continue past warning if it's there
    if Sys.Process("insight").WaitChild('Window("#32770", "Warning", 1)', 3000).Exists:
      click(warning_window_yes()) 
    log("All Slicing complete! Program finished.")
    
    #returns the tab number so we can safely exit out of the config modeler  
    return tab_number
      
    
    
    
  def slice_part_helper(total_number_sliced, save_file_printer, savefile_model_name, 
    savefile_support_name, savefile_model_tip, savefile_support_tip, savefile_slice_height, savefile_part_name):
    
    # Start Slicing Process
    click(insight.green_flag())
     
    # checking to see if the overwrite window pops up or if this is the first item to be sliced
    if not Sys.Process("insight").WaitChild('Window("#32770", "Insight", 1)', 3000).Exists:
      # if not, slices the part and then again so that we can get the naming conventions started
      while Sys.Process("insight").WaitChild('Window("TkTopLevel", "Processing Model", 1)', 1000).Exists:
        wait(ONE_SECOND)
      click(configure_modeler())
      
      # This is the workaround for a resliced item to create a CMB by changing the paths
      click(config_mod_screen.slice_height())
      config_mod_screen.slice_height_dropdown().Keys("[Down]")
      config_mod_screen.slice_height_dropdown().Keys("[Enter]")
      config_mod_screen.modeler_type().Keys("[Tab]" * 3)
      Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
      
      if Sys.Process("insight").WaitChild('Window("#32770", "Warning", 1)', 2000).Exists:
        click(insight.warning_window_yes())
      
      # re setting the slice height to what it should be
      click(configure_modeler())    
      click(config_mod_screen.slice_height())
      config_mod_screen.slice_height_dropdown().Keys("[Up]"*5)
      config_mod_screen.slice_height_dropdown().Keys("[Enter]")
      config_mod_screen.modeler_type().Keys("[Tab]" * 3)  
      Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
      
      click(green_flag())

    
     
    wait(ONE_SECOND)
    while Sys.Process("insight").WaitChild('Window("#32770", "Insight", 1)', 3000).Exists:
      click(overwrite.overwrite_box_no())

    #typing out the new file path
    insight.save_job_window().Keys("[BS][Right]" + savefile_part_name + "_" + save_file_printer +
    "_" + savefile_model_name + "_" + savefile_support_name + "_" + savefile_model_tip + "_" +
    savefile_support_tip + "_" + savefile_slice_height +"slice[Enter]")

    Log.Message(f"Slicing Part Started...")

    # Logging the file name of the saved file
    log(f"NEW FILE SLICED #{total_number_sliced}: {savefile_part_name}_" + save_file_printer + "_" + 
    savefile_model_name + "_" + savefile_support_name + "_" + savefile_model_tip + "_" 
    + savefile_support_tip + "_" + savefile_slice_height +"slice")

    # Waiting for model slicing process to complete
    while Sys.Process("insight").WaitChild('Window("TkTopLevel", "Processing Model", 1)', 1000).Exists:
      wait(ONE_SECOND)
      if Sys.Process("insight").WaitChild('Window("#32770", "Warning", 1)', 1000).Exists:
        click(insight.warning_window_yes())
    wait(ONE_SECOND)
    
    log("Part Slicing Comeplete!!")
    return total_number_sliced + 1
  

    
    
  def configure_modeler_checking():
    """
    
    """
    flag = False
    
            
    # checking model material dropdown
    if Sys.Process("insight").WaitChild('Window("TkTopLevel", "dropdownMaterial", 1)', 1500).Exists:
      log("material drop")
      Sys.Process("insight").Window("TkTopLevel", "dropdownMaterial", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
      #flag = True
      
    # checking slice height
    if Sys.Process("insight").WaitChild('Window("TkTopLevel", "dropsliceHeight", 1)', 1500).Exists:
      log("slice drop")
      Sys.Process("insight").Window("TkTopLevel", "dropsliceHeight", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
      #flag = True
      
    if flag == True:
      click(insight.configure_modeler())
      log("open?")
        
    return
      
    
  def get_printer_move(printer:str):
    return forms.Navigation.move_for_printer(printer)
    
  def get_model_move(printer:str, model_material:str):
    return forms.Navigation.move_for_material(printer, model_material)
    
  def set_printer(printer_name: str):
    printer_number = Looping_Functions.get_printer_move(printer_name)
    
    #insight_to_front()
    click(config_mod_screen.modeler_type())
    
    config_mod_screen.modeler_type_button().Keys("[Enter]")
    config_mod_screen.modeler_dropdown().Keys("[Down]"*printer_number)
    config_mod_screen.modeler_dropdown().Keys("[Enter]")
    
    config_mod_screen.modeler_type().Keys("[Tab][Tab][Tab][Tab][Tab][Tab][Tab][Tab]")
    Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Keys("[Enter]")
    
    Log.Message(f"Printer was set to {printer_name}")
    
    
  def get_move_commands(support_materials, support_name:str, slices,
    slice_name: str, model_tips, model_tip_name:str, support_tips, support_tip_name, tabs):
    """
    This function takes in strings and lists from the excel doc to return a 
    command which will be an int to move down.
    """
    
    support_command = move_helper(support_materials, support_name)
    slice_command = move_helper(slices, slice_name)
    model_tip_command = move_helper(model_tips, model_tip_name)
    support_tip_command = move_helper(support_tips, support_tip_name)
    # since tabs are based on the position of slice in the list, we can use that 
    # as the same position
    tab_command = int(tabs[slice_command])
    
    return support_command, slice_command, model_tip_command, support_tip_command, tab_command
    
    
    
  def move_helper(list, item:str):
    """
    this is a helper function for get_move_commands(), that will take a list and 
    the name of the item in that list and loop through to find out how far down
    the list needs to go.
    """
    number_down = 0
    while item is not list[number_down]:
      number_down += 1
    
    return number_down
    
    
    
  # takes 6 strings
  def type_file_save_info(save_file_printer: str, savefile_model_name: str, 
    savefile_support_name: str, savefile_model_tip: str, 
    savefile_support_tip: str, savefile_slice_height: str):
    """
    Takes in strings and outputs types out the file name based on the unique 
    information from the excel doc.
    """
    return save_job_window().Keys("[Right]_" + save_file_printer + "_" 
    + savefile_model_name + "_" + savefile_support_name + "_" 
    + savefile_model_tip + "_" + savefile_support_tip + "_" 
    + savefile_slice_height +"slice")
  
  
########      END OF FUNCTIONS       #########    
  
  
  
  
    
  
########      START OF BUTTON MAPPINGS       #########  

def configure_modeler():
  return Sys.Process("insight").Window("TkTopLevel", "*",).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 2).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("Button", "", 1)
 
  
def green_flag():
  return Sys.Process("insight").Window("TkTopLevel", "*",).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 2).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 2).Window("Button", "", 5)

def green_flag2():
  return Sys.Process("insight").Window("TkTopLevel", "*", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 2).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 2).Window("Button", "", 5)
  
def overwrite_box():
  return Sys.Process("insight").Window("TkTopLevel", "*", 1) \
  .Window("TkChild", "", 1).Window("TkChild", "", 2).Window("TkChild", "", 3) \
  .Window("TkChild", "", 1).Window("TkChild", "", 1).Window("Button", "", 15)
  

def insight_window():
  return Sys.Process("insight").Window("#32770", "Insight", 1)
  
def save_job_window():
  return Sys.Process("insight").Window("#32770", "Save Job", 1)
  
def warning_window():
  return Sys.Process("insight").Window("#32770", "Warning", 1)
  
def warning_window_yes():
  return Sys.Process("insight").Window("#32770", "Warning", 1).Window("Button", "&Yes", 1)

def warning_window_no():
  return Sys.Process("insight").Window("#32770", "Warning", 1).Window("Button", "&No", 2)

# class for buttons when trying to overwrite saves 
class overwrite():  
  def overwrite_box_yes():
    return Sys.Process("insight").Window("#32770", "Insight", 1).Window("Button", "&Yes", 1)
  
  
  def overwrite_box_no():
    return Sys.Process("insight").Window("#32770", "Insight", 1).Window("Button", "&No", 2)


  def save_in_box():
    return Sys.Process("insight").Window("#32770", "Save Job", 1).Window("ComboBox", "", 1)
  
  
  def job_name_box():  
    return Sys.Process("insight").Window("#32770", "Save Job", 1).Window("ComboBoxEx32", "", 1).Window("ComboBox", "", 1).Window("Edit", "", 1)
  

  def save_button():
    return Sys.Process("insight").Window("#32770", "Save Job", 1).Window("Button", "&Save", 2)
  
  
  def cancel_button():
    return Sys.Process("insight").Window("#32770", "Save Job", 1).Window("Button", "Cancel", 3)  
  
  

     
     
# abreviated path for configure modeler buttons / clicking
def conf_mod_path():
  return Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 3).Window("TkChild", "", 1)

# configure modeler side screen 
class config_mod_screen():

  # radio button All Stratasys modeler types
  def modeler_type_availability():
    return conf_mod_path().Window("TkChild", "", 19).Window("Button", "", 3)
  
  
  # modeler type selection (aka printer gender)
  def modeler_type():
    return conf_mod_path().Window("TkChild", "", 18).Window("TkChild", "", 1)  
  
  # drop down button on right side of 'modeler type'
  def modeler_type_button():
    return conf_mod_path().Window("TkChild", "", 18).Window("Button", "", 1)
  
  # the dropdown menu used for our navigation purposes
  # this was not able to be gotten from the object spy, but the 'map object from screen' button
  def modeler_dropdown():
    return Sys.Process("insight").Window("TkTopLevel", "dropmodelerType", 1).Window("TkChild", "", 1).Window("TkChild", "", 1)
    
  #scroll bar on the dropdown menu
  def modeler_dropdown_scrollbar():
    return NameMapping.Sys.Process("insight").Window("TkTopLevel", "dropmodelerType", 1).Window("TkChild", "", 1).Window("ScrollBar", "", 1)  
  
  
    
  # model material selection
  def model_material():
    return conf_mod_path().Window("TkChild", "", 14).Window("TkChild", "", 1)
  
  # side button to access the drop down menu for model material to be able to use the Up and down commands
  def model_material_button():
    return conf_mod_path().Window("TkChild", "", 14).Window("Button", "", 1)
  
  # model material dropdown menu
  def model_material_dropdown():
    return Sys.Process("insight").Window("TkTopLevel", "dropmodelMaterial", 1).Window("TkChild", "", 1).Window("TkChild", "", 1)
  
    
    
    
  # model material color selection
  def model_material_color():
    return conf_mod_path().Window("TkChild", "", 12).Window("TkChild", "", 1)
  
    
    
  # support material selection
  def support_material():
    return conf_mod_path().Window("TkChild", "", 10).Window("TkChild", "", 1)
  
  # support material button to access the dropdown
  def support_material_button():
    return conf_mod_path().Window("TkChild", "", 10).Window("Button", "", 1)
  
  # support material dropdown menu
  def support_material_dropdown():
    return Sys.Process("insight").Window("TkTopLevel", "dropsupportMaterial", 1).Window("TkChild", "", 1).Window("TkChild", "", 1)
    
    
    
    
  # slice height selection
  def slice_height():
    return conf_mod_path().Window("TkChild", "", 6).Window("TkChild", "", 1)  
  
  # slice height button dropdown to access the dropdown
  def slice_height_button():
    return conf_mod_path().Window("TkChild", "", 6).Window("Button", "", 1)
  
  # slice height dropdown menu
  def slice_height_dropdown():
    return Sys.Process("insight").Window("TkTopLevel", "dropsliceHeight", 1).Window("TkChild", "", 1).Window("TkChild", "", 1)
    
    
    
    
  # model tip selection  
  def model_tip():
    return Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 1).Window("TkChild", "", 1).Window("TkChild", "", 3).Window("TkChild", "", 1).Window("TkChild", "", 4).Window("TkChild", "", 1)
  
  def model_tip_button():
    return Sys.Process("insight").Window("TkTopLevel", "Configure Modeler", 2).Window("TkChild", "", 1).Window("TkChild", "", 3).Window("TkChild", "", 1).Window("TkChild", "", 4).Window("Button", "", 1)
    
  def model_tip_dropdown():
    return Sys.Process("insight").Window("TkTopLevel", "dropmodelDescriptiveTip", 1).Window("TkChild", "", 1).Window("TkChild", "", 1)
      
    
    
    
  # support tip selection  
  def support_tip():
    return conf_mod_path().Window("TkChild", "", 2).Window("TkChild", "", 1)    
    
  def support_tip_button():
      return conf_mod_path().Window("TkChild", "", 2).Window("Button", "", 1)  
    
  # need to map once there is a support tip that can be changed (Currently there is none)
   #def support_tip_dropdown():
     #return conf_mod_path()
      
      
 
     
  # check button to accept
  def check_button():
    return Sys.Process("insight").Window("TkTopLevel", "*", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("Button", "", 3)
  
    
  # x button to cancel
  def x_button():
    return Sys.Process("insight").Window("TkTopLevel", "*", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("TkChild", "", 1).Window("Button", "", 1)
    
########       END OF BUTTON MAPPINGS       #########  