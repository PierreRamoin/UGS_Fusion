# Author-Patrick Rainsberry
# Description-Universal G-Code Sender plugin for Fusion 360
import json
import os
from collections import defaultdict
from dataclasses import dataclass
from os.path import expanduser
import adsk.core
import adsk.fusion
import traceback

# Global Variable to handle command events
handlers = []

# Global Program ID for settings
programID = 'UGS_Fusion'


@dataclass
class Settings:
    ugs_path: str
    ugs_post: str
    ugs_platform: bool
    show_operations: str
    output_folder: str

    def to_json(self):
        return json.dumps(
            self,
            default=lambda o: o.__dict__, 
            sort_keys=True,
            indent=4)


def get_folder():
    # Get user's home directory
    home = expanduser("~")
    home += '/' + programID + '/'

    if not os.path.exists(home):
        os.makedirs(home)

    return home


def get_file_name():
    home = get_folder()
    return home + 'settings.json'


def write_settings(filename, settings):
    with open(filename, 'w') as file:
        file.write(settings.to_json())


def read_settings(filename):
    with open(filename, 'r') as file:
        settings = json.load(file)

    return Settings(**settings)


def get_tool_speed(tool_information, tool_preset_id):
    for preset in tool_information['start-values']['presets']:
        if preset['guid'] == tool_preset_id:
            return preset['n']


def export_file(op_name, settings):
    app = adsk.core.Application.get()
    doc = app.activeDocument
    products = doc.products
    product = products.itemByProductType('CAMProductType')
    cam = adsk.cam.CAM.cast(product)
    to_posts = []
    result_filenames = []
    parent_file_count = defaultdict(int)

    # Currently doesn't handle duplicate in names
    for setup in cam.setups:
        if setup.name == op_name:
            to_posts.append(setup)
        else:
            for folder in setup.folders:
                if folder.name == op_name:
                    to_posts.append(folder)

    for operation in cam.allOperations:
        if operation.name == op_name:
            to_posts.append(operation)

    if op_name == "ALL":
        for operation in cam.allOperations:
            to_posts.append(operation)

    post_config = os.path.join(cam.genericPostFolder, settings.ugs_post)
    units = adsk.cam.PostOutputUnitOptions.DocumentUnitsOutput

    for toPost in to_posts:
        parent_name = toPost.parent.name if toPost.parent is not None else ''
        output_folder_post = f"{settings.output_folder}//{parent_name}"
        filename = (f"{parent_file_count[parent_name]}"
                    f" - {toPost.name}"
                    f" - {toPost.tool.parameters.itemByName('tool_productId').value.value}"
                    f"_{toPost.tool.parameters.itemByName('tool_diameter').value.value:.3f}"
                    f"{toPost.tool.parameters.itemByName('tool_unit').value.value} "
                    f"({int(toPost.tool.parameters.itemByName('tool_spindleSpeed').value.value)} rpm)")

        post_input = adsk.cam.PostProcessInput.create(filename, post_config, output_folder_post, units)
        post_input.isOpenInEditor = False
        cam.postProcess(toPost, post_input)
        parent_file_count[parent_name] += 1

        # Get the resulting filename
        result_filenames += f"{output_folder_post}/{toPost.name}.nc"

    return result_filenames


# Get the current values of the command inputs.
def get_inputs(inputs):
    show_operations_input = inputs.itemById('showOperations')

    settings = Settings(ugs_path=inputs.itemById('UGS_path').text,
                        ugs_post=inputs.itemById('UGS_post').text,
                        ugs_platform=inputs.itemById('UGS_platform').value,
                        output_folder=inputs.itemById('outputFolder').text,
                        show_operations=show_operations_input.selectedItem.name
                        )

    save_settings = inputs.itemById('saveSettings').value
    op_name = None

    # Only attempt to get a value if the user has made a selection
    setup_input = inputs.itemById('setups')
    setup_item = setup_input.selectedItem
    if setup_item:
        setup_name = setup_item.name

    folder_input = inputs.itemById('folders')
    folder_item = folder_input.selectedItem
    if folder_item:
        folder_name = folder_item.name

    operation_input = inputs.itemById('operations')
    operation_item = operation_input.selectedItem
    if operation_item:
        operation_name = operation_item.name

    # Get the name of setup, folder, or operation depending on radio selection
    # This is the operation that will post processed
    if settings.show_operations == 'Setups' and setup_item:
        op_name = setup_name
    elif settings.show_operations == 'Folders':
        op_name = folder_name
    elif settings.show_operations == 'Operations':
        op_name = operation_name
    elif settings.show_operations == 'All Operations':
        op_name = "ALL"

    return op_name, settings, save_settings


# Will update visibility of 3 selection dropdowns based on radio selection
# Also updates radio selection which is only really useful when command is first launched.
def set_dropdown(inputs, show_operations):
    # Get input objects
    setup_input = inputs.itemById('setups')
    folder_input = inputs.itemById('folders')
    operation_input = inputs.itemById('operations')
    show_operations_input = inputs.itemById('showOperations')

    # Set visibility based on appropriate selection from radio list
    if show_operations == 'Setups':
        setup_input.isVisible = True
        folder_input.isVisible = False
        operation_input.isVisible = False
        show_operations_input.listItems[0].isSelected = True
    elif show_operations == 'Folders':
        setup_input.isVisible = False
        folder_input.isVisible = True
        operation_input.isVisible = False
        show_operations_input.listItems[1].isSelected = True
    elif show_operations == 'Operations':
        setup_input.isVisible = False
        folder_input.isVisible = False
        operation_input.isVisible = True
        show_operations_input.listItems[2].isSelected = True
    elif show_operations == 'All Operations':
        setup_input.isVisible = False
        folder_input.isVisible = False
        operation_input.isVisible = False
        show_operations_input.listItems[3].isSelected = True
    else:
        # TODO add error check
        return
    return


# Define the event handler for when the command is executed
class UGSExecutedEventHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            # Get the inputs.
            inputs = args.command.commandInputs
            op_name, settings, save_settings = get_inputs(inputs)

            # Save Settings:
            if save_settings:
                settings_filename = get_file_name()
                write_settings(settings_filename, settings)

            # Export the file and launch UGS
            export_file(op_name, settings)

        except:
            app = adsk.core.Application.get()
            ui = app.userInterface
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Define the event handler for when any input changes.
class UGSInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            # Get inputs and changed inputs
            input_changed = args.input
            inputs = args.inputs

            # Check to see if the post type has changed and show appropriate drop down
            if input_changed.id == 'showOperations':
                show_operations = input_changed.selectedItem.name
                set_dropdown(inputs, show_operations)

        except:
            app = adsk.core.Application.get()
            ui = app.userInterface
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Define the event handler for when the Octoprint command is run by the user.
class UGSCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        ui = []
        try:
            app = adsk.core.Application.get()
            ui = app.userInterface
            doc = app.activeDocument
            products = doc.products
            product = products.itemByProductType('CAMProductType')

            # Check if the document has a CAMProductType. It will not if there are no CAM operations in it.
            if product is None:
                ui.messageBox('There are no CAM operations in the active document')
                return
                # Cast the CAM product to a CAM object (a subtype of product).
            cam = adsk.cam.CAM.cast(product)

            # Setup Handlers and options for command
            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False

            on_execute = UGSExecutedEventHandler()
            cmd.execute.add(on_execute)
            handlers.append(on_execute)

            on_input_changed = UGSInputChangedHandler()
            cmd.inputChanged.add(on_input_changed)
            handlers.append(on_input_changed)

            # Define the inputs.
            inputs = cmd.commandInputs

            # Labels
            inputs.addTextBoxCommandInput('labelText2', '',
                                          '<a href="http://winder.github.io/ugs_website/">Universal Gcode Sender</a></span> A full featured gcode platform used for interfacing with advanced CNC controllers like GRBL and TinyG.',
                                          4, True)
            inputs.addTextBoxCommandInput('labelText3', '', 'Choose the Setup or Operation to send to UGS', 2, True)

            # UGS local path and post information
            ugs_path_input = inputs.addTextBoxCommandInput('UGS_path', 'UGS Path: ', 'Location of UGS', 1, False)
            ugs_post_input = inputs.addTextBoxCommandInput('UGS_post', 'Post to use: ', 'Name of post', 1, False)
            output_folder_input = inputs.addTextBoxCommandInput('outputFolder', 'Output folder: ',
                                                                'Path to output folder', 1, False)

            # Whether using classic or platform
            # TODO Could automate this based on path
            ugs_platform_input = inputs.addBoolValueInput("UGS_platform", 'Using UGS Platform?', True)

            # What to select from?  Setups, Folders, Operations?
            show_operations_input = inputs.addRadioButtonGroupCommandInput("showOperations", 'What to Post?')
            radio_button_items = show_operations_input.listItems
            radio_button_items.add("Setups", False)
            radio_button_items.add("Folders", False)
            radio_button_items.add("Operations", False)
            radio_button_items.add("All Operations", False)

            # Drop down for Setups
            setup_drop_down = inputs.addDropDownCommandInput('setups', 'Select Setup:',
                                                             adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            # Drop down for Folders
            folder_drop_down = inputs.addDropDownCommandInput('folders', 'Select Folder:',
                                                              adsk.core.DropDownStyles.LabeledIconDropDownStyle)
            # Drop down for Operations
            op_drop_down = inputs.addDropDownCommandInput('operations', 'Select Operation:',
                                                          adsk.core.DropDownStyles.LabeledIconDropDownStyle)

            # Populate values in dropdowns based on current document:
            for setup in cam.setups:
                setup_drop_down.listItems.add(setup.name, False)
                for folder in setup.folders:
                    folder_drop_down.listItems.add(folder.name, False)
            for operation in cam.allOperations:
                op_drop_down.listItems.add(operation.name, False)

            # Save user settings
            inputs.addBoolValueInput("saveSettings", 'Save entered settings?', True)

            # Defaults for command dialog
            cmd.commandCategoryName = 'UGS'
            cmd.setDialogInitialSize(500, 300)
            cmd.setDialogMinimumSize(500, 300)
            cmd.okButtonText = 'POST'

            # Check if user has saved settings and update UI to reflect preferences
            settings_file_name = get_file_name()
            if os.path.isfile(settings_file_name):
                settings = read_settings(settings_file_name)

                # Update dialog values
                ugs_path_input.text = settings.ugs_path
                ugs_post_input.text = settings.ugs_post
                ugs_platform_input.value = settings.ugs_platform
                output_folder_input.text = settings.output_folder
                set_dropdown(inputs, settings.show_operations)
            else:
                ugs_post_input.text = 'grbl.cps'
                output_folder_input.text = f'{get_folder()}/output/'
                set_dropdown(inputs, 'Folders')

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def run(context):
    ui = None

    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        if ui.commandDefinitions.itemById('UGSButtonID'):
            ui.commandDefinitions.itemById('UGSButtonID').deleteMe()

        cmd_defs = ui.commandDefinitions

        # Create a button command definition for the comamnd button.  This
        # is also used to display the disclaimer dialog.
        tooltip = '<div style=\'font-family:"Calibri";color:#B33D19; padding-top:-20px;\'><span style=\'font-size:20px;\'><b>winder.github.io/ugs_website</b></span></div>Universal Gcode Sender'
        ugs_button_def = cmd_defs.addButtonDefinition('UGSButtonID', 'Post to UGS', tooltip, './/Resources')
        on_ugs_created = UGSCreatedEventHandler()
        ugs_button_def.commandCreated.add(on_ugs_created)
        handlers.append(on_ugs_created)

        # Find the "ADD-INS" panel for the solid and the surface workspaces.
        solid_panel = ui.allToolbarPanels.itemById('CAMActionPanel')

        # Add a button for the "Request Quotes" command into both panels.
        solid_panel.controls.addCommand(ugs_button_def, '', False)
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface

        if ui.commandDefinitions.itemById('UGSButtonID'):
            ui.commandDefinitions.itemById('UGSButtonID').deleteMe()

        # Find the controls in the solid and surface panels and delete them.
        cam_panel = ui.allToolbarPanels.itemById('CAMActionPanel')
        cntrl = cam_panel.controls.itemById('UGSButtonID')
        if cntrl:
            cntrl.deleteMe()


    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
