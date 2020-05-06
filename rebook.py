#!/usr/bin/python

from threading import Thread
from tkinter import *
from tkinter.ttk import *
import asyncio
import glob
import json
import os
import subprocess as sub
import tkinter.filedialog
import tkinter.messagebox
import tkinter.scrolledtext as scrltxt


def check_page_nums(input_pages):
    page_num_list = re.split(',|-|o|e', input_pages)

    for page_num in page_num_list:
        if len(page_num) > 0 and not page_num.isdigit():
            return FALSE
    return TRUE

# GUI construction

root = Tk()

# screen_width = root.winfo_screenwidth()
# screen_height = root.winfo_screenheight()

# root.resizable(FALSE, FALSE)

root.title('rebook')

# main tab

baseTab = Notebook(root)

conversionTab = Frame(baseTab)
baseTab.add(conversionTab, text='Conversion')

logTab = Frame(baseTab)
baseTab.add(logTab, text='Log')

baseTab.pack(expand=1, fill='both')

# menu

menu_bar = Menu(root)
root['menu'] = menu_bar

def on_command_about_box_cb():
    theMessage = \
        '''rebook

TclTk GUI for k2pdfopt by Pu Wang

The source code can be found at:
http://github.com/pwang7/rebook/rebook.py'''

    tkinter.messagebox.showinfo(message=theMessage)

menu_file = Menu(menu_bar)
menu_bar.add_cascade(menu=menu_file, label='File')
menu_file.add_command(label='About', command=on_command_about_box_cb)

# root.createcommand('tkAboutDialog', on_command_about_box_cb)
# root.createcommand('::tk::mac::ShowHelp', on_command_about_box_cb)

# global variables and functions

k2pdfopt_path = './k2pdfopt'
custom_preset_file_path = 'rebook_preset.json'

def check_k2pdfopt_path_exists():
    if not os.path.exists(k2pdfopt_path):
        tkinter.messagebox.showerror(
            message='Failed to find k2pdfopt, ' +
            'please put it under the same directory ' +
            'as rebook and then restart.'
        )
        quit()

k2pdfopt_cmd_args = {}

def update_cmd_arg_entry_strvar():
    global strvarCmdArgs

    strvarCmdArgs.set(generate_cmd_arg_str())

def add_or_update_one_cmd_arg(arg_key, arg_value):
    global k2pdfopt_cmd_args

    k2pdfopt_cmd_args[arg_key] = arg_value

def remove_one_cmd_arg(arg_key):
    global k2pdfopt_cmd_args

    previous = k2pdfopt_cmd_args.pop(arg_key, None)

    return previous

def load_custom_preset():
    global strvarOutputFilePath

    if os.path.exists(custom_preset_file_path):
        with open(custom_preset_file_path) as preset_file:
            dict_to_load = json.load(preset_file)

            if dict_to_load:
                log_string('Load Preset: ' + str(dict_to_load))

                initialize_vars(dict_to_load)

                return TRUE

    return FALSE

def log_string(str_line):
    global stdoutText

    log_content = str_line.strip()
    if len(log_content) > 0:
        stdoutText.config(state=NORMAL)

        print('=== ' + log_content)  # TODO: remove print

        stdoutText.insert(END, log_content + '\n')
        stdoutText.config(state=DISABLED)

def clear_logs():
    stdoutText.config(state=NORMAL)
    stdoutText.delete(1.0, END)
    stdoutText.config(state=DISABLED)

def initialize_vars(dict_vars):
    global k2pdfopt_cmd_args

    for k, v in dict_vars.items():
        for i in range(len(v)):
            arg_var_map[k][i].set(v[i])

    k2pdfopt_cmd_args = {}

    for cb_func in arg_cb_map.values():
        if cb_func is not None:
            cb_func()
    # must be after loading preset values
    update_cmd_arg_entry_strvar()

def restore_default_values():
    clear_logs()

    remove_preview_img_and_clear_canvas()

    for sv in string_var_list:
        sv.set('')

    for bv in bool_var_list:
        bv.set(FALSE)

    for b in combo_box_list:
        b.current(0)

    initialize_vars(default_var_map)

def start_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_forever()

thread_loop = asyncio.get_event_loop()
run_loop_thread = Thread(
    target=start_loop,
    args=(thread_loop,),
    daemon=TRUE,
)
run_loop_thread.start()

background_process = None
background_future = None

device_argument_map = {
    0: 'k2',
    1: 'dx',
    2: 'kpw',
    3: 'kp2',
    4: 'kp3',
    5: 'kv',
    6: 'ko2',
    7: 'pb2',
    8: 'nookst',
    9: 'kbt',
    10: 'kbg',
    11: 'kghd',
    12: 'kghdfs',
    13: 'kbm',
    14: 'kba',
    15: 'kbhd',
    16: 'kbh2o',
    17: 'kbh2ofs',
    18: 'kao',
    19: 'nex7',
    20: None,
}

device_choice_map = {
    0: 'Kindle 1-5',
    1: 'Kindle DX',
    2: 'Kindle Paperwhite',
    3: 'Kindle Paperwhite 2',
    4: 'Kindle Paperwhite 3',
    5: 'Kindle Voyage/PW3+/Oasis',
    6: 'Kindle Oasis 2',
    7: 'Pocketbook Basic 2',
    8: 'Nook Simple Touch',
    9: 'Kobo Touch',
    10: 'Kobo Glo',
    11: 'Kobo Glo HD',
    12: 'Kobo Glo HD Full Screen',
    13: 'Kobo Mini',
    14: 'Kobo Aura',
    15: 'Kobo Aura HD',
    16: 'Kobo H2O',
    17: 'Kobo H2O Full Screen',
    18: 'Kobo Aura One',
    19: 'Nexus 7',
    20: 'Other (specify width & height)',
}

mode_argument_map = {
    0: 'def',
    1: 'copy',
    2: 'fp',
    3: 'fw',
    4: '2col',
    5: 'tm',
    6: 'crop',
    7: 'concat',
}

mode_choice_map = {
    0: 'Default',
    1: 'Copy',
    2: 'Fit Page',
    3: 'Fit Width',
    4: '2 Columns',
    5: 'Trim Margins',
    6: 'Crop',
    7: 'Concat',
}

unit_argument_map = {
    0: 'in',
    1: 'cm',
    2: 's',
    3: 't',
    4: 'p',
    5: 'x',
}

unit_choice_map = {
    0: 'Inches',
    1: 'Centimeters',
    2: 'Source Page Size',
    3: 'Trimmed Source Region Size',
    4: 'Pixels',
    5: 'Relative to the OCR Text Layer',
}

def generate_cmd_arg_str():
    must_have_args = '-a- -ui- -x'

    device_arg = k2pdfopt_cmd_args.pop(device_arg_name, None)
    if device_arg is None:
        width_arg = k2pdfopt_cmd_args.pop(width_arg_name)
        height_arg = k2pdfopt_cmd_args.pop(height_arg_name)

    mode_arg = k2pdfopt_cmd_args.pop(conversion_mode_arg_name)

    arg_list = [mode_arg] + list(k2pdfopt_cmd_args.values())

    k2pdfopt_cmd_args[conversion_mode_arg_name] = mode_arg
    if device_arg is not None:
        arg_list.append(device_arg)

        k2pdfopt_cmd_args[device_arg_name] = device_arg
    else:
        arg_list.append(width_arg)
        arg_list.append(height_arg)

        k2pdfopt_cmd_args[width_arg_name] = width_arg
        k2pdfopt_cmd_args[height_arg_name] = height_arg

    arg_list.append(must_have_args)

    log_string('Generate Argument List: ' + str(arg_list))

    cmd_arg_str = ' '.join(arg_list)

    return cmd_arg_str

def convert_pdf_file(output_arg):
    check_k2pdfopt_path_exists()

    async def async_run_cmd_and_log(exec_cmd):
        global background_process

        executed = exec_cmd.strip()

        def log_bytes(log_btyes):
            log_string(log_btyes.decode('utf-8'))

        log_string(executed)

        p = await asyncio.create_subprocess_shell(
            executed,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        background_process = p

        while True:
            line = await p.stdout.readline()

            log_bytes(line)

            if not line:
                break
            if line == '' and p.returncode is not None:
                break

    input_pdf_path = strvarFilePath.get().strip()
    if ' ' in input_pdf_path:
        # in case the file name contains space
        input_pdf_path = '\"' + input_pdf_path + '\"'

    executed = ' '.join([
        k2pdfopt_path,
        input_pdf_path,
        output_arg,
        generate_cmd_arg_str(),
    ])
    future = asyncio.run_coroutine_threadsafe(
        async_run_cmd_and_log(executed),
        thread_loop,
    )
    return future

def check_pdf_conversion_done():
    if (background_future is None) or (background_future.done()):
        if ((background_process is None) or
                (background_process.returncode is not None)):
            return TRUE

    tkinter.messagebox.showerror(
        message='Background Conversion Not Finished Yet! Please Wait.',
    )
    return FALSE

# conversion tab
conversion_tab_left_part_column_num = 0
conversion_tab_left_part_row_num = -1

# required inputs
conversion_tab_left_part_row_num += 1

device_arg_name = '-dev'  # -dev <name>
width_arg_name = '-w'  # -w <width>[in|cm|s|t|p]
height_arg_name = '-h'  # -h <height>[in|cm|s|t|p|x]
conversion_mode_arg_name = '-mode'  # -mode <mode>
output_path_arg_name = '-o'  # -o <namefmt>
output_pdf_suffix = '-output.pdf'
screen_unit_prefix = '-screen_unit'

strvarFilePath = StringVar()
strvarDevice = StringVar()
strvarConversionMode = StringVar()
strvarScreenUnit = StringVar()
strvarScreenWidth = StringVar()
strvarScreenHeight = StringVar()

requiredInputFrame = Labelframe(conversionTab, text='Required Inputs')
requiredInputFrame.grid(
    column=conversion_tab_left_part_column_num,
    row=conversion_tab_left_part_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

required_frame_row_num = 0

inputPathEntry = Entry(
    requiredInputFrame,
    state='readonly',
    textvariable=strvarFilePath,
)
inputPathEntry.grid(
    column=0,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

def on_command_open_pdf_file_cb():
    formats = [('PDF files', '*.pdf')]

    filename = tkinter.filedialog.askopenfilename(
        parent=root,
        filetypes=formats,
        title='Choose a file',
    )

    if filename is not None and len(filename.strip()) > 0:
        strvarFilePath.set(filename)
        (base_path, file_ext) = os.path.splitext(filename)
        strvarOutputFilePath.set(base_path + output_pdf_suffix)

openButton = Button(
    requiredInputFrame,
    text='Open PDF File',
    command=on_command_open_pdf_file_cb,
)
openButton.grid(
    column=1,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

required_frame_row_num += 1

deviceText = Label(requiredInputFrame, text='Device:')
deviceText.grid(
    column=0,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

def update_device_unit_width_height():
    if deviceComboBox.current() != 20:  # non-other type
        deviceType = device_argument_map[deviceComboBox.current()]
        arg = device_arg_name + ' ' + deviceType
        add_or_update_one_cmd_arg(device_arg_name, arg)

        remove_one_cmd_arg(width_arg_name)
        remove_one_cmd_arg(height_arg_name)
    else:
        screen_unit = unit_argument_map[unitComboBox.current()]

        width_arg = (
            width_arg_name + ' ' +
            strvarScreenWidth.get().strip() + screen_unit
        )
        add_or_update_one_cmd_arg(width_arg_name, width_arg)

        height_arg = (
            height_arg_name + ' ' +
            strvarScreenHeight.get().strip() + screen_unit
        )
        add_or_update_one_cmd_arg(height_arg_name, height_arg)

        remove_one_cmd_arg(device_arg_name)

def on_bind_event_device_unit_cb(e=None):
    update_device_unit_width_height()

deviceComboBox = Combobox(
    requiredInputFrame,
    state='readonly',
    textvariable=strvarDevice,
)
deviceComboBox['values'] = list(device_choice_map.values())
deviceComboBox.current(0)
deviceComboBox.bind('<<ComboboxSelected>>', on_bind_event_device_unit_cb)
deviceComboBox.grid(
    column=1,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

required_frame_row_num += 1

unitTextLabel = Label(requiredInputFrame, text='Unit:')
unitTextLabel.grid(
    column=0,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

unitComboBox = Combobox(
    requiredInputFrame,
    state='readonly',
    textvariable=strvarScreenUnit,
)
unitComboBox['values'] = list(unit_choice_map.values())
unitComboBox.current(0)
unitComboBox.bind('<<ComboboxSelected>>', on_bind_event_device_unit_cb)
unitComboBox.grid(
    column=1,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

required_frame_row_num += 1

def on_command_width_height_cb():
    update_device_unit_width_height()

widthTextLabel = Label(requiredInputFrame, text='Width:')
widthTextLabel.grid(
    column=0,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

widthSpinBox = Spinbox(
    requiredInputFrame,
    from_=0,
    to=10000,
    increment=0.1,
    state='readonly',
    textvariable=strvarScreenWidth,
    command=on_command_width_height_cb,
)
widthSpinBox.grid(
    column=1,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

required_frame_row_num += 1

heightTextLabel = Label(requiredInputFrame, text='Height:')
heightTextLabel.grid(
    column=0,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

heightSpinBox = Spinbox(
    requiredInputFrame,
    from_=0,
    to=10000,
    increment=0.1,
    state='readonly',
    textvariable=strvarScreenHeight,
    command=on_command_width_height_cb,
)
heightSpinBox.grid(
    column=1,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

required_frame_row_num += 1

def on_bind_event_mode_cb(e=None):
    conversionMode = mode_argument_map[modeComboBox.current()]
    arg = (
        conversion_mode_arg_name + ' ' +
        conversionMode
    )
    add_or_update_one_cmd_arg(conversion_mode_arg_name, arg)

modeTextLabel = Label(requiredInputFrame, text='Conversion Mode:')
modeTextLabel.grid(
    column=0,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

modeComboBox = Combobox(
    requiredInputFrame,
    state='readonly',
    textvariable=strvarConversionMode,
)
modeComboBox['values'] = list(mode_choice_map.values())
modeComboBox.current(0)
modeComboBox.bind('<<ComboboxSelected>>', on_bind_event_mode_cb)
modeComboBox.grid(
    column=1,
    row=required_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

# info frame
conversion_tab_left_part_row_num += 1

strvarOutputFilePath = StringVar()
strvarCmdArgs = StringVar()

infoFrame = Labelframe(conversionTab, text='Related Info')
infoFrame.grid(
    column=conversion_tab_left_part_column_num,
    row=conversion_tab_left_part_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

def on_command_save_cb():
    with open(custom_preset_file_path, 'w') as preset_file:
        dict_to_save = {}
        for k, v in arg_var_map.items():
            dict_to_save[k] = [var.get() for var in v]

        json.dump(dict_to_save, preset_file)

saveTextLabel = Label(infoFrame, text='Save Current Setting as Preset:')
saveTextLabel.grid(column=0, row=0, sticky=N+W, pady=0, padx=5)

saveButton = Button(infoFrame, text='Save', command=on_command_save_cb)
saveButton.grid(column=1, row=0, sticky=N+W, pady=0, padx=5)

outputTextLabel = Label(infoFrame, text='Output Pdf File Path:')
outputTextLabel.grid(column=0, row=1, sticky=N+W, pady=0, padx=5)

outputPathEntry = Entry(
    infoFrame,
    state='readonly',
    textvariable=strvarOutputFilePath,
)
outputPathEntry.grid(column=1, row=1, sticky=N+W, pady=0, padx=5)

cmdArgTextLabel = Label(infoFrame, text='Command-line Options:')
cmdArgTextLabel.grid(column=0, row=2, sticky=N+W, pady=0, padx=5)

def on_bind_event_cmd_args_cb(e=None):
    update_cmd_arg_entry_strvar()

cmdArgEntry = Entry(
    infoFrame,
    state='readonly',
    textvariable=strvarCmdArgs,
)
cmdArgEntry.bind('<Button-1>', on_bind_event_cmd_args_cb)
cmdArgEntry.grid(column=1, row=2, sticky=N+W, pady=0, padx=5)

# parameters
conversion_tab_left_part_row_num += 1

column_num_arg_name = '-col'  # -col <maxcol>
resolution_multiplier_arg_name = '-dr'  # -dr <value>
crop_margin_arg_name = '-cbox'  # -cbox[<pagelist>|u|-]
dpi_arg_name = '-dpi'  # -dpi <dpival>
page_num_arg_name = '-p'  # -p <pagelist>

isColumnNum = BooleanVar()
isResolutionMultipler = BooleanVar()
isCropMargin = BooleanVar()
isDPI = BooleanVar()

strvarColumnNum = StringVar()
strvarResolutionMultiplier = StringVar()
strvarCropPageRange = StringVar()
strvarLeftMargin = StringVar()
strvarRightMargin = StringVar()
strvarTopMargin = StringVar()
strvarBottomMargin = StringVar()
strvarDPI = StringVar()
strvarPageNums = StringVar()

paraFrame = Labelframe(conversionTab, text='Parameters')
paraFrame.grid(
    column=conversion_tab_left_part_column_num,
    row=conversion_tab_left_part_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num = 0

def on_command_column_num_cb():
    if isColumnNum.get():
        arg = (
            column_num_arg_name + ' ' +
            strvarColumnNum.get().strip()
        )
        add_or_update_one_cmd_arg(column_num_arg_name, arg)
    else:
        remove_one_cmd_arg(column_num_arg_name)

colNumCheckButton = Checkbutton(
    paraFrame,
    text='Max Columns:',
    variable=isColumnNum,
    command=on_command_column_num_cb,
)
colNumCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

colNumSpinBox = Spinbox(
    paraFrame,
    from_=1,
    to=10,
    increment=1,
    state='readonly',
    textvariable=strvarColumnNum,
    command=on_command_column_num_cb,
)
colNumSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def on_command_resolution_multipler_cb():
    if isResolutionMultipler.get():
        arg = (
            resolution_multiplier_arg_name + ' ' +
            strvarResolutionMultiplier.get().strip()
        )
        add_or_update_one_cmd_arg(resolution_multiplier_arg_name, arg)
    else:
        remove_one_cmd_arg(resolution_multiplier_arg_name)

resolutionCheckButton = Checkbutton(
    paraFrame,
    text='Document Resolution Factor:',
    variable=isResolutionMultipler,
    command=on_command_resolution_multipler_cb,
)
resolutionCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

resolutionSpinBox = Spinbox(
    paraFrame,
    from_=0.1,
    to=10.0,
    increment=0.1,
    state='readonly',
    textvariable=strvarResolutionMultiplier,
    command=on_command_resolution_multipler_cb,
)
resolutionSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def on_command_and_validate_crop_margin_cb():
    if (len(strvarCropPageRange.get().strip()) > 0 and
            not check_page_nums(strvarCropPageRange.get().strip())):
        remove_one_cmd_arg(crop_margin_arg_name)
        strvarCropPageRange.set('')

        tkinter.messagebox.showerror(
            message='Invalide Crop Page Range! Should be like 2-5e,3-7o,9-',
        )

        return FALSE

    if isCropMargin.get():
        page_range_arg = strvarCropPageRange.get().strip()
        margin_args = [
            strvarLeftMargin.get(),
            strvarTopMargin.get(),
            strvarRightMargin.get(),
            strvarBottomMargin.get(),
        ]
        arg = (
            # no space between -cbox and page range
            crop_margin_arg_name + page_range_arg + ' '
            'in,'.join(map(str.strip, margin_args)) + 'in'
        )
        add_or_update_one_cmd_arg(crop_margin_arg_name, arg)
    else:
        remove_one_cmd_arg(crop_margin_arg_name)

marginCheckButton = Checkbutton(
    paraFrame,
    text='Crop Margins (in):',
    variable=isCropMargin,
    command=on_command_and_validate_crop_margin_cb,
)
marginCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

cropPageRangeTextLabel = Label(paraFrame, text='Page Range:')
cropPageRangeTextLabel.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

cropPageRangeEntry = Entry(
    paraFrame,
    textvariable=strvarCropPageRange,
    validate='focusout',
    validatecommand=on_command_and_validate_crop_margin_cb,
)
cropPageRangeEntry.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

leftMarginTextLabel = Label(paraFrame, text='Left Margin:')
leftMarginTextLabel.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

leftMarginSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=100,
    increment=0.01,
    state='readonly',
    textvariable=strvarLeftMargin,
    command=on_command_and_validate_crop_margin_cb,
)
leftMarginSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

rightMarginTextLabel = Label(
    paraFrame,
    text='Right Margin:',
)
rightMarginTextLabel.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

rightMarginSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=100,
    increment=0.01,
    state='readonly',
    textvariable=strvarRightMargin,
    command=on_command_and_validate_crop_margin_cb,
)
rightMarginSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

topMarginTextLabel = Label(paraFrame, text='Top Margin:')
topMarginTextLabel.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

topMarginSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=100,
    increment=0.01,
    state='readonly',
    textvariable=strvarTopMargin,
    command=on_command_and_validate_crop_margin_cb,
)
topMarginSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

bottomMarginTextLabel = Label(paraFrame, text='Bottom Margin:')
bottomMarginTextLabel.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

bottomMarginSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=100,
    increment=0.01,
    state='readonly',
    textvariable=strvarBottomMargin,
    command=on_command_and_validate_crop_margin_cb,
)
bottomMarginSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def on_command_dpi_cb():
    if isDPI.get():
        arg = dpi_arg_name + ' ' + strvarDPI.get().strip()
        add_or_update_one_cmd_arg(dpi_arg_name, arg)
    else:
        remove_one_cmd_arg(dpi_arg_name)

dpiCheckButton = Checkbutton(
    paraFrame,
    text='DPI:',
    variable=isDPI,
    command=on_command_dpi_cb,
)
dpiCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

dpiSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=1000,
    increment=1,
    state='readonly',
    textvariable=strvarDPI,
    command=on_command_dpi_cb,
)
dpiSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def validate_and_update_page_nums():
    if (len(strvarPageNums.get().strip()) > 0 and
            not check_page_nums(strvarPageNums.get().strip())):
        remove_one_cmd_arg(page_num_arg_name)
        strvarPageNums.set('')

        tkinter.messagebox.showerror(
            message='Invalide Page Argument! Should be like 2-5e,3-7o,9-',
        )

        return FALSE

    if len(strvarPageNums.get().strip()) > 0:
        arg = page_num_arg_name + ' ' + strvarPageNums.get().strip()
        add_or_update_one_cmd_arg(page_num_arg_name, arg)
    else:
        remove_one_cmd_arg(page_num_arg_name)

    return TRUE

pageNumTextLabel = Label(
    paraFrame,
    text='Pages to Convert:',
)
pageNumTextLabel.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

def on_validate_page_nums_cb():
    validate_and_update_page_nums()

pageNumEntry = Entry(
    paraFrame,
    textvariable=strvarPageNums,
    validate='focusout',
    validatecommand=on_validate_page_nums_cb,
)
pageNumEntry.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

fixed_font_size_arg_name = '-fs'  # -fs 0/-fs <font size>[+]
ocr_arg_name = '-ocr'  # -ocr-/-ocr t
ocr_cpu_arg_name = '-nt'  # -nt -50/-nt <percentage>
landscape_arg_name = '-ls'  # -ls[-][pagelist]
linebreak_arg_name = '-ws'  # -ws <spacing>

isFixedFontSize = BooleanVar()
isOCR = BooleanVar()
isLandscape = BooleanVar()
isLinebreak = BooleanVar()  # -ws 0.01~10

strvarFixedFontSize = StringVar()
strvarOcrCpuPercentage = StringVar()
strvarLandscapePages = StringVar()  # 1,3,5-10
strvarLinebreakSpace = StringVar()

# checkbox with value options
para_frame_row_num += 1

def on_command_fixed_font_size_cb():
    if isFixedFontSize.get():
        arg = (
            fixed_font_size_arg_name + ' ' +
            strvarFixedFontSize.get().strip()
        )
        add_or_update_one_cmd_arg(
            fixed_font_size_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(fixed_font_size_arg_name)

fontSizeCheckButton = Checkbutton(
    paraFrame,
    text='Fixed Output Font Size:',
    variable=isFixedFontSize,
    command=on_command_fixed_font_size_cb,
)
fontSizeCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

fontSizeSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=100,
    increment=1,
    state='readonly',
    textvariable=strvarFixedFontSize,
    command=on_command_fixed_font_size_cb,
)
fontSizeSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def on_command_ocr_and_cpu_cb():
    if isOCR.get():
        # ocr conflicts with native pdf
        isNativePdf.set(FALSE)
        remove_one_cmd_arg(native_pdf_arg_name)

        ocr_arg = ocr_arg_name
        add_or_update_one_cmd_arg(ocr_arg_name, ocr_arg)

        # negtive integer means percentage
        ocr_cpu_arg = (
            ocr_cpu_arg_name + '-' +
            strvarOcrCpuPercentage.get().strip()
        )
        add_or_update_one_cmd_arg(
            ocr_cpu_arg_name,
            ocr_cpu_arg,
        )
    else:
        remove_one_cmd_arg(ocr_arg_name)
        remove_one_cmd_arg(ocr_cpu_arg_name)

ocrCheckButton = Checkbutton(
    paraFrame,
    text='OCR (Tesseract) and CPU %:',
    variable=isOCR,
    command=on_command_ocr_and_cpu_cb,
)
ocrCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

ocrCpuSpinBox = Spinbox(
    paraFrame,
    from_=0,
    to=100,
    increment=1,
    state='readonly',
    textvariable=strvarOcrCpuPercentage,
    command=on_command_ocr_and_cpu_cb,
)
ocrCpuSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def on_command_and_validate_landscape_cb():
    if (len(strvarLandscapePages.get().strip()) > 0 and
            not check_page_nums(strvarLandscapePages.get().strip())):
        remove_one_cmd_arg(landscape_arg_name)
        strvarLandscapePages.set('')

        tkinter.messagebox.showerror(
            message='Invalide Landscape Page Argument!',
        )

        return FALSE

    if isLandscape.get():
        arg = '-ls'
        if len(strvarLandscapePages.get().strip()) > 0:
            # no space between -ls and page numbers
            arg += strvarLandscapePages.get()

        add_or_update_one_cmd_arg(landscape_arg_name, arg.strip())
    else:
        remove_one_cmd_arg(landscape_arg_name)

    return TRUE

landscapeCheckButton = Checkbutton(
    paraFrame,
    text='Output in Landscape:',
    variable=isLandscape,
    command=on_command_and_validate_landscape_cb,
)
landscapeCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

landscapePageNumEntry = Entry(
    paraFrame,
    textvariable=strvarLandscapePages,
    validate='focusout',
    validatecommand=on_command_and_validate_landscape_cb,
)
landscapePageNumEntry.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

para_frame_row_num += 1

def on_command_line_break_cb():
    if isLinebreak.get():
        arg = (
            linebreak_arg_name + ' ' +
            strvarLinebreakSpace.get().strip()
        )
        add_or_update_one_cmd_arg(linebreak_arg_name, arg)
    else:
        remove_one_cmd_arg(linebreak_arg_name)

lineBreakCheckButton = Checkbutton(
    paraFrame,
    text='Smart Line Breaks:',
    variable=isLinebreak,
    command=on_command_line_break_cb,
)
lineBreakCheckButton.grid(
    column=0,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

lineBreakSpinBox = Spinbox(
    paraFrame,
    from_=0.01,
    to=2.00,
    increment=0.01,
    state='readonly',
    textvariable=strvarLinebreakSpace,
    command=on_command_line_break_cb,
)
lineBreakSpinBox.grid(
    column=1,
    row=para_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

# right side of conversion tab
conversion_tab_right_part_column_num = 1
conversion_tab_right_part_row_num = -1

# options
conversion_tab_right_part_row_num += 1

auto_straignten_arg_name = '-as'  # -as-/-as
break_page_avoid_overlap_arg_name = '-bp'  # -bp-/-bp
color_output_arg_name = '-c'  # -c-/-c
native_pdf_arg_name = '-n'  # -n-/-n
right_to_left_arg_name = '-r'  # -r-/-r
post_gs_arg_name = '-ppgs'  # -ppgs-/-ppgs
marked_source_arg_name = '-sm'  # -sm-/-sm
reflow_text_arg_name = '-wrap'  # -wrap+/-wrap-
erase_vertical_line_arg_name = '-evl'  # -evl 0/-evl 1
erase_horizontal_line_arg_name = '-ehl'  # -ehl 0/-ehl 1
fast_preview_arg_name = '-rt'  # -rt /-rt 0
ign_small_defects_arg_name = '-de'  # -de 1.0/-de 1.5
auto_crop_arg_name = '-ac'  # -ac-/-ac

isAutoStraighten = BooleanVar()
isBreakPage = BooleanVar()
isColorOutput = BooleanVar()
isNativePdf = BooleanVar()
isRight2Left = BooleanVar()
isPostGs = BooleanVar()
isMarkedSrc = BooleanVar()
isReflowText = BooleanVar()
isEraseVerticalLine = BooleanVar()
isEraseHorizontalLine = BooleanVar()
isFastPreview = BooleanVar()
isAvoidOverlap = BooleanVar()
isIgnSmallDefects = BooleanVar()
isAutoCrop = BooleanVar()

optionFrame = Labelframe(
    conversionTab,
    text='Options',
)
optionFrame.grid(
    column=conversion_tab_right_part_column_num,
    row=conversion_tab_right_part_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

option_frame_left_part_col_num = 0
option_frame_row_num = 0

def on_command_auto_straighten_cb():
    if isAutoStraighten.get():
        arg = auto_straignten_arg_name
        add_or_update_one_cmd_arg(auto_straignten_arg_name, arg)
    else:
        remove_one_cmd_arg(auto_straignten_arg_name)

opt1 = Checkbutton(
    optionFrame,
    text='Autostraighten',
    variable=isAutoStraighten,
    command=on_command_auto_straighten_cb,
)
opt1.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_break_page_cb():
    if isBreakPage.get():
        # break page conflicts with avoid overlap since they are both -bp flag
        isAvoidOverlap.set(FALSE)
        remove_one_cmd_arg(break_page_avoid_overlap_arg_name)

        arg = break_page_avoid_overlap_arg_name
        add_or_update_one_cmd_arg(
            break_page_avoid_overlap_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(break_page_avoid_overlap_arg_name)

opt2 = Checkbutton(
    optionFrame,
    text='Break After Each Source Page',
    variable=isBreakPage,
    command=on_command_break_page_cb,
)
opt2.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_color_output_cb():
    if isColorOutput.get():
        arg = color_output_arg_name
        add_or_update_one_cmd_arg(color_output_arg_name, arg)
    else:
        remove_one_cmd_arg(color_output_arg_name)

opt3 = Checkbutton(
    optionFrame,
    text='Color Output',
    variable=isColorOutput,
    command=on_command_color_output_cb,
)
opt3.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_native_pdf_cb():
    if isNativePdf.get():
        # native pdf conflicts with ocr and reflow text
        isOCR.set(FALSE)
        remove_one_cmd_arg(ocr_arg_name)
        remove_one_cmd_arg(ocr_cpu_arg_name)

        isReflowText.set(FALSE)
        remove_one_cmd_arg(reflow_text_arg_name)

        arg = native_pdf_arg_name
        add_or_update_one_cmd_arg(native_pdf_arg_name, arg)
    else:
        remove_one_cmd_arg(native_pdf_arg_name)

opt4 = Checkbutton(
    optionFrame,
    text='Native PDF Output',
    variable=isNativePdf,
    command=on_command_native_pdf_cb,
)
opt4.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_right_to_left_cb():
    if isRight2Left.get():
        arg = right_to_left_arg_name
        add_or_update_one_cmd_arg(right_to_left_arg_name, arg)
    else:
        remove_one_cmd_arg(right_to_left_arg_name)

opt5 = Checkbutton(
    optionFrame,
    text='Right-to-Left Text',
    variable=isRight2Left,
    command=on_command_right_to_left_cb,
)
opt5.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_post_gs_cb():
    if isPostGs.get():
        arg = post_gs_arg_name
        add_or_update_one_cmd_arg(post_gs_arg_name, arg)
    else:
        remove_one_cmd_arg(post_gs_arg_name)

opt6 = Checkbutton(
    optionFrame,
    text='Post Process w/GhostScript',
    variable=isPostGs,
    command=on_command_post_gs_cb,
)
opt6.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_marked_src_cb():
    if isMarkedSrc.get():
        arg = marked_source_arg_name
        add_or_update_one_cmd_arg(
            marked_source_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(marked_source_arg_name)

opt7 = Checkbutton(
    optionFrame,
    text='Generate Marked-up Source',
    variable=isMarkedSrc,
    command=on_command_marked_src_cb,
)
opt7.grid(
    column=option_frame_left_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

option_frace_right_part_col_num = 1
option_frame_row_num = 0

def on_command_reflow_text_cb():
    if isReflowText.get():
        # reflow text conflicts with native pdf
        isNativePdf.set(FALSE)
        remove_one_cmd_arg(native_pdf_arg_name)

        arg = reflow_text_arg_name + '+'
        add_or_update_one_cmd_arg(
            reflow_text_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(reflow_text_arg_name)

opt8 = Checkbutton(
    optionFrame,
    text='Re-flow Text',
    variable=isReflowText,
    command=on_command_reflow_text_cb,
)
opt8.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_erase_vertical_line_cb():
    if isEraseVerticalLine.get():
        arg = erase_vertical_line_arg_name + ' 1'
        add_or_update_one_cmd_arg(
            erase_vertical_line_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(erase_vertical_line_arg_name)

opt9 = Checkbutton(
    optionFrame,
    text='Erase Vertical Lines',
    variable=isEraseVerticalLine,
    command=on_command_erase_vertical_line_cb,
)
opt9.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_erase_horizontal_line_cb():
    if isEraseHorizontalLine.get():
        arg = erase_horizontal_line_arg_name + ' 1'
        add_or_update_one_cmd_arg(
            erase_horizontal_line_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(erase_horizontal_line_arg_name)

opt14 = Checkbutton(
    optionFrame,
    text='Erase Horizontal Lines',
    variable=isEraseHorizontalLine,
    command=on_command_erase_horizontal_line_cb,
)
opt14.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_fast_preview_cb():
    if isFastPreview.get():
        arg = fast_preview_arg_name + ' 0'
        add_or_update_one_cmd_arg(fast_preview_arg_name, arg)
    else:
        remove_one_cmd_arg(fast_preview_arg_name)

opt10 = Checkbutton(
    optionFrame,
    text='Fast Preview',
    variable=isFastPreview,
    command=on_command_fast_preview_cb,
)
opt10.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_avoid_text_selection_overlap_cb():
    if isAvoidOverlap.get():
        # avoid overlap conflicts with break page since they are both -bp flag
        isBreakPage.set(FALSE)
        remove_one_cmd_arg(break_page_avoid_overlap_arg_name)

        arg = break_page_avoid_overlap_arg_name + ' m'
        add_or_update_one_cmd_arg(
            break_page_avoid_overlap_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(break_page_avoid_overlap_arg_name)

opt11 = Checkbutton(
    optionFrame,
    text='Avoid Text Selection Overlap',
    variable=isAvoidOverlap,
    command=on_command_avoid_text_selection_overlap_cb,
)
opt11.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_ign_small_defect_cb():
    if isIgnSmallDefects.get():
        arg = (
            ign_small_defects_arg_name + ' 1.5'
        )
        add_or_update_one_cmd_arg(
            ign_small_defects_arg_name,
            arg,
        )
    else:
        remove_one_cmd_arg(ign_small_defects_arg_name)

opt12 = Checkbutton(
    optionFrame,
    text='Ignore Small Defects',
    variable=isIgnSmallDefects,
    command=on_command_ign_small_defect_cb,
)
opt12.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

def on_command_auto_crop_cb():
    if isAutoCrop.get():
        arg = auto_crop_arg_name
        add_or_update_one_cmd_arg(auto_crop_arg_name, arg)
    else:
        remove_one_cmd_arg(auto_crop_arg_name)

opt13 = Checkbutton(
    optionFrame,
    text='Auto-Crop',
    variable=isAutoCrop,
    command=on_command_auto_crop_cb,
)
opt13.grid(
    column=option_frace_right_part_col_num,
    row=option_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
option_frame_row_num += 1

# preview frame
conversion_tab_right_part_row_num += 1

preview_output_arg_name = '-bmp'
preview_image_path = './k2pdfopt_out.png'
current_preview_page_index = 1

# global variable to hold opened preview image to prevent gc collecting it
preview_img = None
canvas_image_tag = None


strvarCurrentPreviewPageNum = StringVar()

def remove_preview_img_and_clear_canvas():
    global canvas_image_tag
    global strvarCurrentPreviewPageNum

    if os.path.exists(preview_image_path):

        os.remove(preview_image_path)

    previewImageCanvas.delete(ALL)
    canvas_image_tag = None

def load_image_to_canvas(photo_img, canvas):
        load_image_to_canvas(preview_img, previewImageCanvas)

def load_preview_image(img_path, preview_page_index):
    # PhotoImage must be global var to prevent gc collect it
    global preview_img
    global previewImageCanvas

    if os.path.exists(img_path):
        preview_img = PhotoImage(file=img_path)

        canvas_image_tag = previewImageCanvas.create_image(
            (0, 0),
            anchor=NW,
            image=preview_img,
            tags='preview',
        )

        (left_pos, top_pos, right_pos, bottom_pos) = (
            0,
            0,
            preview_img.width(),
            preview_img.height(),
        )
        previewImageCanvas.config(
            scrollregion=(left_pos, top_pos, right_pos, bottom_pos),
        )
        # canvas.scale('preview', 0, 0, 0.1, 0.1)
        strvarCurrentPreviewPageNum.set('Page: ' + str(preview_page_index))
    else:
        strvarCurrentPreviewPageNum.set('No Page ' + str(preview_page_index))

def generate_one_preview_image(preview_page_index):
    global background_future

    if not check_pdf_conversion_done():
        return

    if not os.path.exists(strvarFilePath.get().strip()):
        tkinter.messagebox.showerror(
            message=(
                "Failed to Find Input PDF File to convert for Preview: %s"
                %
                strvarFilePath.get().strip()
            ),
        )
        return

    remove_preview_img_and_clear_canvas()

    (base_path, file_ext) = os.path.splitext(strvarFilePath.get().strip())

    output_arg = ' '.join([preview_output_arg_name, str(preview_page_index)])

    background_future = convert_pdf_file(output_arg)

    strvarCurrentPreviewPageNum.set('Preview Generating...')

    def preview_image_future_cb(bgf):
        load_preview_image(preview_image_path, preview_page_index)
        log_string(
            "Preview generation for page %d finished" %
            preview_page_index
        )

    background_future.add_done_callback(preview_image_future_cb)

previewFrame = Labelframe(conversionTab, text='Preview & Convert')
previewFrame.grid(
    column=conversion_tab_right_part_column_num,
    row=conversion_tab_right_part_row_num,
    rowspan=3,
    sticky=N+S+E+W,
    pady=0,
    padx=5,
)

preview_frame_row_num = 0

def on_command_restore_default_cb():
    restore_default_values()

resetButton = Button(
    previewFrame,
    text='Reset Default',
    command=on_command_restore_default_cb,
)
resetButton.grid(
    column=0,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

def on_command_abort_conversion_cb():
    global background_future
    global background_process

    if background_future is not None:
        background_future.cancel()

    if (background_process is not None and
            background_process.returncode is None):
        background_process.terminate()

cancelButton = Button(
    previewFrame,
    text='Abort',
    command=on_command_abort_conversion_cb,
)
cancelButton.grid(
    column=1,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

def on_command_convert_pdf_cb():
    if not check_pdf_conversion_done():
        return

    global background_future

    pdf_output_arg = output_path_arg_name + ' %s' + output_pdf_suffix
    background_future = convert_pdf_file(pdf_output_arg)

convertButton = Button(
    previewFrame,
    text='Convert',
    command=on_command_convert_pdf_cb,
)
convertButton.grid(
    column=2,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

preview_frame_row_num += 1

currentPreviewPageNumEntry = Entry(
    previewFrame,
    state='readonly',
    textvariable=strvarCurrentPreviewPageNum)
currentPreviewPageNumEntry.grid(
    column=0,
    row=preview_frame_row_num,
    columnspan=2,
    sticky=N+W,
    pady=0,
    padx=5,
)

def on_command_ten_page_up_cb():
    global current_preview_page_index
    current_preview_page_index -= 10
    if current_preview_page_index < 1:
        current_preview_page_index = 1
    generate_one_preview_image(current_preview_page_index)

previewButton = Button(
    previewFrame,
    text='Preview',
    command=on_command_ten_page_up_cb,
)
previewButton.grid(
    column=2,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)

preview_frame_column_num = 0
preview_frame_row_num += 1

firstButton = Button(
    previewFrame,
    text='<<',
    command=on_command_ten_page_up_cb,
)
firstButton.grid(
    column=preview_frame_column_num,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
preview_frame_column_num += 1

def on_command_page_up_cb():
    global current_preview_page_index
    if current_preview_page_index > 1:
        current_preview_page_index -= 1
    generate_one_preview_image(current_preview_page_index)

previousButton = Button(
    previewFrame,
    text='<',
    command=on_command_page_up_cb,
)
previousButton.grid(
    column=preview_frame_column_num,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
preview_frame_column_num += 1

def on_command_page_down_cb():
    global current_preview_page_index
    current_preview_page_index += 1
    generate_one_preview_image(current_preview_page_index)

nextButton = Button(
    previewFrame,
    text='>',
    command=on_command_page_down_cb,
)
nextButton.grid(
    column=preview_frame_column_num,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
preview_frame_column_num += 1

def on_command_ten_page_down_cb():
    global current_preview_page_index
    current_preview_page_index += 10
    generate_one_preview_image(current_preview_page_index)

lastButton = Button(
    previewFrame,
    text='>>',
    command=on_command_ten_page_down_cb,
)
lastButton.grid(
    column=preview_frame_column_num,
    row=preview_frame_row_num,
    sticky=N+W,
    pady=0,
    padx=5,
)
preview_frame_column_num += 1

preview_frame_row_num += 1

xScrollBar = Scrollbar(previewFrame, orient=HORIZONTAL)
xScrollBar.grid(
    column=0,
    row=preview_frame_row_num+1,
    columnspan=preview_frame_column_num,
    sticky=E+W,
)

yScrollBar = Scrollbar(previewFrame)
yScrollBar.grid(
    column=preview_frame_column_num,
    row=preview_frame_row_num,
    sticky=N+S,
)

previewImageCanvas = Canvas(
    previewFrame,
    bd=0,
    xscrollcommand=xScrollBar.set,
    yscrollcommand=yScrollBar.set,
)
previewImageCanvas.grid(
    column=0,
    row=preview_frame_row_num,
    columnspan=preview_frame_column_num,
    sticky=N+S+E+W,
)

xScrollBar.config(command=previewImageCanvas.xview)
yScrollBar.config(command=previewImageCanvas.yview)

conversionTab.columnconfigure(
    conversion_tab_right_part_column_num,
    weight=1,
)
conversionTab.rowconfigure(
    conversion_tab_right_part_row_num,
    weight=1,
)
previewFrame.columnconfigure(0, weight=1)
previewFrame.rowconfigure(preview_frame_row_num, weight=1)

def yscroll_canvas(event):
    previewImageCanvas.yview_scroll(-1 * event.delta, 'units')

def xscroll_canvas(event):
    previewImageCanvas.xview_scroll(-1 * event.delta, 'units')

previewImageCanvas.bind('<MouseWheel>', yscroll_canvas)
previewImageCanvas.bind("<Shift-MouseWheel>", xscroll_canvas)

preview_frame_row_num += 1

# collect all vars

bool_var_list = [
    isColumnNum,
    isResolutionMultipler,
    isCropMargin,
    isDPI,
    isFixedFontSize,
    isOCR,
    isLandscape,
    isLinebreak,

    isAutoStraighten,
    isBreakPage,
    isColorOutput,
    isNativePdf,
    isRight2Left,
    isPostGs,
    isMarkedSrc,
    isReflowText,
    isEraseVerticalLine,
    isEraseHorizontalLine,
    isFastPreview,
    isAvoidOverlap,
    isIgnSmallDefects,
    isAutoCrop,
]

string_var_list = [
    strvarFilePath,
    strvarDevice,
    strvarConversionMode,
    strvarScreenUnit,
    strvarScreenWidth,
    strvarScreenHeight,

    strvarOutputFilePath,
    strvarCmdArgs,

    strvarColumnNum,
    strvarResolutionMultiplier,
    strvarCropPageRange,
    strvarLeftMargin,
    strvarRightMargin,
    strvarTopMargin,
    strvarBottomMargin,
    strvarDPI,
    strvarPageNums,

    strvarFixedFontSize,
    strvarOcrCpuPercentage,
    strvarLandscapePages,
    strvarLinebreakSpace,

    strvarCurrentPreviewPageNum,
]

combo_box_list = [
    deviceComboBox,
    modeComboBox,
    unitComboBox,
]

entry_list = [
    inputPathEntry,
    outputPathEntry,
    cmdArgEntry,
    pageNumEntry,
    landscapePageNumEntry,
    currentPreviewPageNumEntry,
]

default_var_map = {
    device_arg_name:                    ['Kindle 1-5'],
    screen_unit_prefix:                 ['Pixels'],
    width_arg_name:                     ['560'],
    height_arg_name:                    ['735'],
    conversion_mode_arg_name:           ['Default'],
    output_path_arg_name:               [''],

    column_num_arg_name:                [FALSE, '2'],
    resolution_multiplier_arg_name:     [FALSE, '1.0'],
    crop_margin_arg_name:               [
                                            FALSE,
                                            '',
                                            '0.00',
                                            '0.00',
                                            '0.00',
                                            '0.00',
                                        ],
    dpi_arg_name:                       [FALSE, '167'],
    page_num_arg_name:                  [''],
    fixed_font_size_arg_name:           [FALSE, '12'],
    ocr_arg_name:                       [FALSE, '50'],
    ocr_cpu_arg_name:                   [FALSE, '50'],
    landscape_arg_name:                 [FALSE, ''],
    linebreak_arg_name:                 [TRUE, '0.200'],

    auto_straignten_arg_name:           [FALSE],
    break_page_avoid_overlap_arg_name:  [FALSE, FALSE],
    color_output_arg_name:              [FALSE],
    native_pdf_arg_name:                [FALSE],
    right_to_left_arg_name:             [FALSE],
    post_gs_arg_name:                   [FALSE],
    marked_source_arg_name:             [FALSE],
    reflow_text_arg_name:               [TRUE],
    erase_vertical_line_arg_name:       [FALSE],
    erase_horizontal_line_arg_name:     [FALSE],
    fast_preview_arg_name:              [TRUE],
    ign_small_defects_arg_name:         [FALSE],
    auto_crop_arg_name:                 [FALSE],

    preview_output_arg_name:            []
}

arg_var_map = {
    device_arg_name:                    [strvarDevice],
    screen_unit_prefix:                 [strvarScreenUnit],
    width_arg_name:                     [strvarScreenWidth],
    height_arg_name:                    [strvarScreenHeight],
    conversion_mode_arg_name:           [strvarConversionMode],
    output_path_arg_name:               [strvarOutputFilePath],

    column_num_arg_name:                [
                                            isColumnNum,
                                            strvarColumnNum,
                                        ],
    resolution_multiplier_arg_name:     [
                                            isResolutionMultipler,
                                            strvarResolutionMultiplier,
                                        ],
    crop_margin_arg_name:               [
                                            isCropMargin,
                                            strvarCropPageRange,
                                            strvarLeftMargin,
                                            strvarTopMargin,
                                            strvarRightMargin,
                                            strvarBottomMargin,
                                        ],
    dpi_arg_name:                       [
                                            isDPI,
                                            strvarDPI,
                                        ],
    page_num_arg_name:                  [
                                            strvarPageNums,
                                        ],

    fixed_font_size_arg_name:           [
                                            isFixedFontSize,
                                            strvarFixedFontSize,
                                        ],
    ocr_arg_name:                       [
                                            isOCR,
                                            strvarOcrCpuPercentage,
                                        ],
    ocr_cpu_arg_name:                   [
                                            isOCR,
                                            strvarOcrCpuPercentage,
                                        ],
    landscape_arg_name:                 [
                                            isLandscape,
                                            strvarLandscapePages,
                                        ],
    linebreak_arg_name:                 [
                                            isLinebreak,
                                            strvarLinebreakSpace,
                                        ],

    auto_straignten_arg_name:           [isAutoStraighten],
    break_page_avoid_overlap_arg_name:  [isBreakPage, isAvoidOverlap],
    color_output_arg_name:              [isColorOutput],
    native_pdf_arg_name:                [isNativePdf],
    right_to_left_arg_name:             [isRight2Left],
    post_gs_arg_name:                   [isPostGs],
    marked_source_arg_name:             [isMarkedSrc],
    reflow_text_arg_name:               [isReflowText],
    erase_vertical_line_arg_name:       [isEraseVerticalLine],
    erase_horizontal_line_arg_name:     [isEraseHorizontalLine],
    fast_preview_arg_name:              [isFastPreview],
    # break_page_avoid_overlap_arg_name:  []
    ign_small_defects_arg_name:         [isIgnSmallDefects],
    auto_crop_arg_name:                 [isAutoCrop],
    preview_output_arg_name:            []
}

arg_cb_map = {
    device_arg_name:                   on_bind_event_device_unit_cb,
    width_arg_name:                    on_command_width_height_cb,
    height_arg_name:                   on_command_width_height_cb,
    conversion_mode_arg_name:          on_bind_event_mode_cb,
    output_path_arg_name:              None,

    column_num_arg_name:               on_command_column_num_cb,
    resolution_multiplier_arg_name:    on_command_resolution_multipler_cb,
    crop_margin_arg_name:              on_command_and_validate_crop_margin_cb,
    dpi_arg_name:                      on_command_dpi_cb,
    page_num_arg_name:                 on_validate_page_nums_cb,

    fixed_font_size_arg_name:          on_command_fixed_font_size_cb,
    ocr_arg_name:                      on_command_ocr_and_cpu_cb,
    ocr_cpu_arg_name:                  on_command_ocr_and_cpu_cb,
    landscape_arg_name:                on_command_and_validate_landscape_cb,
    linebreak_arg_name:                on_command_line_break_cb,

    auto_straignten_arg_name:          on_command_auto_straighten_cb,
    break_page_avoid_overlap_arg_name: on_command_break_page_cb,
    color_output_arg_name:             on_command_color_output_cb,
    native_pdf_arg_name:               on_command_native_pdf_cb,
    right_to_left_arg_name:            on_command_right_to_left_cb,
    post_gs_arg_name:                  on_command_post_gs_cb,
    marked_source_arg_name:            on_command_marked_src_cb,
    reflow_text_arg_name:              on_command_reflow_text_cb,
    erase_vertical_line_arg_name:      on_command_erase_vertical_line_cb,
    erase_horizontal_line_arg_name:    on_command_erase_horizontal_line_cb,
    fast_preview_arg_name:             on_command_fast_preview_cb,
    ign_small_defects_arg_name:        on_command_ign_small_defect_cb,
    auto_crop_arg_name:                on_command_auto_crop_cb,

    preview_output_arg_name:           None
}

# k2pdfopt stdout

stdoutFrame = Labelframe(logTab, text='k2pdfopt STDOUT:')
stdoutFrame.pack(expand=1, fill='both')

def on_command_clear_log_cb():
    clear_logs()

clearButton = Button(
    stdoutFrame,
    text='Clear',
    command=on_command_clear_log_cb)
clearButton.grid(
    column=0,
    row=0,
    sticky=N+W,
    pady=0,
    padx=5,
)

stdoutText = scrltxt.ScrolledText(
    stdoutFrame,
    state=DISABLED, wrap='word',
)
stdoutText.grid(column=0, row=1, sticky=N+S+E+W)
stdoutFrame.columnconfigure(0, weight=1)
stdoutFrame.rowconfigure(1, weight=1)
# stdoutText.pack(expand=1, fill='both')

# initialization

def initialize():
    check_k2pdfopt_path_exists()

    if not load_custom_preset():
        restore_default_values()

    pwd = os.getcwd()
    log_string('Current directory: ' + pwd)

initialize()

# start TclTk loop

root.mainloop()
