import json
import shutil
import sys
from tkinter import *
from tkinter import filedialog, ttk, messagebox
import os
from openpyxl.utils.exceptions import InvalidFileException
import reg_qpcr
import chip_qpcr
import send_email
from logger import log_exceptions


# Determine if the application is running as a script or bundled executable
if getattr(sys, 'frozen', False):
    # The application is bundled
    basedir = sys._MEIPASS
else:
    # The application is running as a script
    basedir = os.path.dirname(os.path.abspath(__file__))


def open_file(output_filename):
    file_path = filedialog.askopenfilename(
        filetypes=[('Microsoft Excel Worksheet', ('*.xls', '*.xlsx')), ('All Files', '*.*')])

    if file_path:
        # this gets the full path of your selected file_tkinter.TclError: couldn't open "C:\Users\jason\AppData\Local\Temp\_MEI248642\cog-icon.png": no such file or directory
        filename = file_path
        # this is only selecting the name with file extension
        filename = os.path.abspath(filename)
        # then create a new label with the filename
        output_filename.set(filename)


SENDER_EMAIL = 'qpcrcalculator@gmail.com'
APP_PASSWORD = 'nqlkimcvxryclbwt'


class App(Tk):
    @log_exceptions
    def __init__(self):
        super().__init__()
        self.iconbitmap(os.path.join(basedir, 'pipette.ico'))
        self.current_row_value = 3
        self.created_widgets = []
        self.entry_widgets = []
        self.existing_labels = set()
        self.hkgs = []
        self.combinations = []
        self.emails = []
        self.additional_row_counter = 0
        self.file_uploaded = False
        self.cont_file_uploaded = False

        # Setting up initial things
        self.title("ΔΔCт Calculator")
        self.bind("<Button-1>", self.on_window_click)
        self.state("zoom")
        self.outerframe = Frame(self)
        self.outerframe.pack(expand=True, anchor="center", fill=BOTH)
        self.minsize(width=500, height=400)

        self.canvas = Canvas(self.outerframe)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # Add vertical scrollbar to the canvas
        vertical_scrollbar = Scrollbar(self.outerframe, orient="vertical", command=self.canvas.yview)
        vertical_scrollbar.grid(row=0, column=1, sticky='ns')
        self.canvas.config(yscrollcommand=vertical_scrollbar.set)

        # Add horizontal scrollbar to the canvas
        horizontal_scrollbar = Scrollbar(self.outerframe, orient="horizontal", command=self.canvas.xview)
        horizontal_scrollbar.grid(row=1, column=0, sticky='ew')
        self.canvas.config(xscrollcommand=horizontal_scrollbar.set)

        self.canvas_width = 1600
        self.canvas_height = 900
        self.canvas.configure(scrollregion=(0, 0, self.canvas_width, self.canvas_height))

        # Configure the frame's grid
        self.outerframe.grid_rowconfigure(0, weight=1)
        self.outerframe.grid_columnconfigure(0, weight=1)

        # Create a header frame
        self.header_frame = Frame(self.canvas)
        self.header_frame_window = self.canvas.create_window((0, 0), window=self.header_frame, anchor="center")
        canvas_center_x = self.canvas_width / 2
        canvas_center_y = self.canvas_height / 2
        self.canvas.coords(self.header_frame_window, canvas_center_x, canvas_center_y)

        self.header_frame.bind("<Configure>", self.update_scrollregion)
        self.canvas.bind("<Configure>", self.on_canvas_resize)

        self.inner_frame = Frame(self.header_frame)
        self.inner_frame.grid(column=1, row=0)

        # Bind the mouse wheel event
        self.bind_all("<MouseWheel>", self.on_mousewheel)

        # Load the cog icon
        self.cog_icon = PhotoImage(file=os.path.join(basedir, 'cog-icon.png'))

        # Settings button
        self.settings_button = Button(self.header_frame, image=self.cog_icon, command=self.open_settings, relief=FLAT)
        self.settings_button.grid(column=2, row=0, sticky="n")

        # Fake button to even stuff out
        self.dummy_label = Label(self.header_frame, width=self.settings_button.cget("width"))
        self.dummy_label.grid(column=0, row=0)

        # Mode
        self.modeframe = Frame(self.inner_frame)
        self.modeframe.grid(column=0, row=0)
        self.mode = StringVar()
        self.mode_selection = ttk.Combobox(self.modeframe, textvariable=self.mode,
                                           values=["qPCR ΔΔCт", "qPCR ΔΔCт - Continuous", "ChIP qPCR"],
                                           state="readonly")
        self.mode_selection.current(0)
        self.mode_selection.grid(column=0, row=0, pady=20)

        # Upload Files
        self.filename = StringVar()
        self.filename.set("Select qPCR Results File")
        self.output_filename_cont = StringVar()
        self.output_filename_cont.set("")
        self.data = []
        self.targets = []

        self.choose_file_frame = Frame(self.inner_frame)
        self.choose_file_frame.grid(column=0, row=1)
        file_label = Label(self.choose_file_frame, textvariable=self.filename)
        file_label.grid(column=0, row=1, pady=20, sticky=W + E)

        self.current_max_col = 2
        upload_button = Button(self.choose_file_frame, text="Choose Export File", command=lambda: (self.handle_upload()), width=16)
        upload_button.grid(column=1, row=1, padx=50, sticky="W")

        # Existing Files set up for continuous
        self.existing_filename = StringVar()
        self.existing_filename.set("Select Existing File")

        self.mode_selection.bind('<<ComboboxSelected>>', lambda _: self.switch_modes())

    def default_settings(self):
        """Return the default settings."""
        return {
            "REF_MODE": "Exclude Other References",
            "DECIMAL_PLACES": 100,
            "BATCH_TEXT_COLOR": "ff003822",
            "BATCH_FILL_COLOR": "ffedffdc",
            "AVERAGE_BATCHES_NAME": "Average (Do not delete this if you want continuous input)",
            "PERCENTAGE_INPUT": 1,
            "SAMPLE_COLUMN": ["sample", "sample name", "name"],
            "WELL_COLUMN": ["well"],
            "GENE_COLUMN": ["gene", "target", "target name", "detector", "primer/probe", "primer", "probe"],
            "CT_COLUMN": ["ct", "cq", "cт"]
        }

    def get_settings(self):
        """Load settings from JSON file. If the file doesn't exist, return default settings."""
        try:
            with open("settings.json", "r") as file:
                return json.load(file)
        except FileNotFoundError:
            return self.default_settings()

    def save_to_file(self, settings):
        """Save settings to JSON file."""
        with open("settings.json", "w") as file:
            json.dump(settings, file)

    def open_settings(self):

        widget_width = 31

        def single_setting(text, json_key, label_row, label_column):
            # Setting 1 Label and Entry
            Label(container, text=text).grid(row=label_row, column=label_column, pady=15, padx=10)
            entry_setting = Entry(container, width=widget_width)
            entry_setting.insert(0, current_settings.get(json_key))
            entry_setting.grid(row=label_row, column=label_column + 1, pady=15, padx=10)

            return json_key, entry_setting

        def multi_settings(text, json_key, label_row, label_column):
            Label(container, text=text).grid(row=label_row, column=label_column, pady=(10, 0), padx=10, sticky="N")

            frame = Frame(container, padx=15, pady=10)
            frame.grid(row=label_row, column=label_column + 1)
            preset_listbox = Listbox(frame, height=5, width=widget_width)
            for preset in current_settings.get(json_key):
                preset_listbox.insert(END, preset)
            preset_listbox.grid(row=1, column=0, padx=10)

            # Button to remove a selected preset
            remove_button = Button(frame, text="Remove Selected Preset",
                                   command=lambda: self.remove_preset(preset_listbox), width=widget_width - 5)
            remove_button.grid(row=2, column=0, padx=10)

            small_frame = Frame(frame)
            small_frame.grid(row=3, column=0)

            # Entry and button to add new presets
            new_preset_entry = Entry(small_frame)
            new_preset_entry.grid(row=0, column=0, pady=10)
            add_button = Button(small_frame, text="Add Preset",
                                command=lambda: self.add_preset(preset_listbox, new_preset_entry))
            add_button.grid(row=0, column=1, pady=10)

            return json_key, preset_listbox

        def dropdown_option(options, text, json_key, label_row, label_column):
            """Create a dropdown with the given options and initial value."""
            Label(container, text=text).grid(row=label_row, column=label_column, pady=15, padx=10)
            dropdown_var = StringVar()
            dropdown_combobox = ttk.Combobox(container, textvariable=dropdown_var, values=options, state="readonly",
                                             width=widget_width - 3)
            dropdown_combobox.grid(row=label_row, column=label_column + 1)
            return json_key, dropdown_var

        current_settings = self.get_settings()
        single_settings_list = []
        multi_settings_list = []

        settings_window = Toplevel(self)
        settings_window.title("Settings")

        # container = Frame(settings_window, padx=20, pady=20)
        # container.pack(expand=True, fill=BOTH)

        # Create a canvas and a scrollbar
        canvas = Canvas(settings_window)
        scrollbar = Scrollbar(settings_window, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create the container frame to hold all widgets, and add it to the canvas
        container = Frame(canvas, padx=20, pady=20)
        canvas.create_window((0, 0), window=container, anchor="nw")

        # Function to update the scrollregion of the canvas
        def onFrameConfigure(canvas):
            canvas.configure(scrollregion=canvas.bbox("all"))

        container.bind("<Configure>", lambda event, canvas=canvas: onFrameConfigure(canvas))

        # Function to handle mouse scroll
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"  # Prevents default event propagation to parent widgets

        # Bind the scroll event to the canvas
        settings_window.bind("<MouseWheel>", _on_mousewheel)

        # Cleanup function to unbind the mousewheel event when the window is closed
        def on_close():
            settings_window.unbind("<MouseWheel>")
            settings_window.destroy()

        # Bind the cleanup function to the window close event
        settings_window.protocol("WM_DELETE_WINDOW", on_close)

        # Place the canvas and the scrollbar in the window
        canvas.pack(side="left", fill=BOTH, expand=True)
        scrollbar.pack(side="right", fill="y")

        # multiple ref mode
        # Dropdown for setting
        self.ref_mode_var = dropdown_option(options=["Exclude Other References", "Include Other References"],
                                            json_key="REF_MODE", text="Multiple References", label_row=0,
                                            label_column=0)

        # decimal_places
        decimal_places = single_setting(text="Decimal Places", json_key="DECIMAL_PLACES", label_row=1, label_column=0)

        # batch text color
        batch_text_color = single_setting(text="Batch Text Color", json_key="BATCH_TEXT_COLOR", label_row=2,
                                          label_column=0)

        # batch fill color
        batch_fill_color = single_setting(text="Batch Fill Color", json_key="BATCH_FILL_COLOR", label_row=3,
                                          label_column=0)

        # Average batches name
        average_batches_name = single_setting(text="Average Batches Name", json_key="AVERAGE_BATCHES_NAME", label_row=4,
                                              label_column=0)

        # chip percentage input
        percent_input = single_setting(text="ChIP qPCR % Input", json_key="PERCENTAGE_INPUT", label_row=5,
                                       label_column=0)

        # Sample col
        sample_col = multi_settings(text="Sample Column Name", json_key="SAMPLE_COLUMN", label_row=6, label_column=0)

        # well col
        well_col = multi_settings(text="Well Column Name", json_key="WELL_COLUMN", label_row=7, label_column=0)

        # gene col
        gene_col = multi_settings(text="Gene Column Name", json_key="GENE_COLUMN", label_row=8, label_column=0)

        # ct column
        ct_col = multi_settings(text="CT Column Name", json_key="CT_COLUMN", label_row=9, label_column=0)

        single_settings_list.extend(
            [decimal_places, batch_text_color, batch_fill_color, average_batches_name,
             percent_input])
        multi_settings_list.extend([sample_col, well_col, gene_col, ct_col])

        def save_settings():
            current_settings[self.ref_mode_var[0]] = self.ref_mode_var[1].get()

            for setting in single_settings_list:
                current_settings[setting[0]] = setting[1].get()

            for setting in multi_settings_list:
                current_settings[setting[0]] = list(setting[1].get(0, END))

            self.save_to_file(current_settings)
            messagebox.showinfo("Info", "Settings Saved")
            settings_window.destroy()

        def reset_settings():
            default_settings = self.default_settings()
            self.ref_mode_var[1].set(default_settings[self.ref_mode_var[0]])

            for setting in single_settings_list:
                setting[1].delete(0, END)
                setting[1].insert(0, default_settings[setting[0]])

            for setting in multi_settings_list:
                setting[1].delete(0, END)
                setting[1].insert(0, *default_settings[setting[0]])

            self.save_to_file(self.default_settings())
            messagebox.showinfo("Info", "Settings Saved")
            settings_window.destroy()

        btn_frame = Frame(container)
        btn_frame.grid(row=10, column=0, columnspan=2, pady=(20, 0))

        Button(btn_frame, text="Reset to Default", command=reset_settings, width=15).pack(side=LEFT, padx=(0, 10))
        Button(btn_frame, text="Save", command=save_settings, width=15).pack(side=LEFT)

        self.ref_mode_var[1].set(self.get_settings().get("REF_MODE"))

    def add_preset(self, listbox, entry):
        """Add a new preset from the entry to the listbox."""
        new_preset = entry.get()
        if new_preset and new_preset not in listbox.get(0, END):  # Only add if non-empty and unique
            listbox.insert(END, new_preset)

    def remove_preset(self, listbox):
        """Remove the selected preset from the listbox."""
        try:
            selected_index = listbox.curselection()[0]
            listbox.delete(selected_index)
        except IndexError:
            pass  # No item was selected

    def show_error(self, text):
        messagebox.showerror("Error", text)

    def update_scrollregion(self, event):
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def on_canvas_resize(self, event=None):
        self.canvas.update_idletasks()
        # Adjust the width of the inner frame to match the canvas
        self.inner_frame.config(width=self.canvas.winfo_width())
        self.inner_frame.config(height=self.canvas.winfo_height())
        # Update the scroll region
        self.update_scrollregion(event)

        canvas_center_x = self.canvas.winfo_width() / 2
        canvas_center_y = self.canvas.winfo_height() / 2
        self.canvas.coords(self.header_frame_window, canvas_center_x, canvas_center_y)

        # Dynamically adjust the canvas's scrollregion to the new size of the header_frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


    def on_mousewheel(self, event):
        self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def clear_widgets(self, exceptions=None):
        if exceptions is None:
            exceptions = []
        for widget in self.created_widgets:
            if widget not in exceptions:
                widget.destroy()
        self.created_widgets = []
        self.existing_labels = set()

    def current_row(self, add_row=0):
        row = self.current_row_value
        self.current_row_value += add_row
        return row

    def acquire_data(self, file_name):
        if self.mode.get() in ["qPCR ΔΔCт", "qPCR ΔΔCт - Continuous"]:
            self.data = reg_qpcr.get_data(file_name)
        elif self.mode.get() == "ChIP qPCR":
            self.data = chip_qpcr.get_data(file_name, what_sort=self.orientation.get())
            # need to ask for orientation

    def get_targets(self, data_list):
        if self.mode.get() == "ChIP qPCR":
            targets_dup = data_list[1]
        else:
            targets_dup = data_list[2]
        counter = 0
        frame = Frame(self.targets_container, bg="white", padx=5, pady=5)
        frame.pack(side="top", anchor="center")

        for num in range(len(targets_dup)):
            if counter % 2 == 0 and counter > 0:
                self.current_row_value += 1

            var = IntVar()
            checkbox = Checkbutton(frame, text=targets_dup[num], variable=var, onvalue=1, offvalue=0, bg="white")
            self.targets.append(targets_dup[num])
            checkbox.grid(column=(num % 2), row=self.current_row_value, padx=50, sticky=W)
            self.created_widgets.append(checkbox)
            self.hkgs.append([var, targets_dup[num]])
            counter += 1

        self.current_row_value += 1

    def setup_targets_area(self):
        label = Label(self.inner_frame, text="Error", font=('Helvetica', 12, 'bold'))
        if self.mode.get() in ["qPCR ΔΔCт", "qPCR ΔΔCт - Continuous"]:
            label = Label(self.inner_frame, text="Choose your reference genes", font=('Helvetica', 12, 'bold'))
        elif self.mode.get() == "ChIP qPCR":
            label = Label(self.inner_frame, text="Choose your reference samples e.g. CTRL Input",
                          font=('Helvetica', 12, 'bold'))
        label.grid(column=0, row=self.current_row(1), pady=(20, 0))
        self.created_widgets.append(label)
        self.targets_container = Frame(self.inner_frame, bg="white", bd=2, relief=SUNKEN, padx=10, pady=5)
        self.targets_container.grid(row=self.current_row(1), column=0, sticky=W + E + N + S, padx=10, pady=10)
        self.created_widgets.append(self.targets_container)

    def clear_content(self, event):
        # Function to clear the content of the combobox
        event.widget.set('')

    def add_combos(self, data_list, first=False):
        fc_targets = data_list[1]
        row_num = self.current_row(0)
        frame = Frame(self.combo_frame, bg="lightgrey", padx=5, pady=5)
        frame.grid(sticky=W + E)
        pair = []
        for num in range(2):
            var = StringVar()
            drop = ttk.Combobox(frame, textvariable=var, state="readonly", values=fc_targets)
            drop.grid(column=num, row=row_num)
            pair.append(var)
            drop.bind("<BackSpace>", self.clear_content)  # Bind the <BackSpace> key event
            self.created_widgets.append(drop)
        self.combinations.append(pair)

        close_button = Button(frame, text="✖", command=lambda: self.destroy_combo(frame, pair), bg="lightgrey",
                              fg="red", relief=FLAT, font=("Courier", 10))
        close_button.grid(row=self.current_row(1), column=2, sticky=E)

    def destroy_combo(self, frame, pair_list):
        self.combinations.remove(pair_list)
        frame.destroy()

    def clear_all_comboboxes(self, comboboxes):
        """Clear the content of all comboboxes in the list."""
        for combo in comboboxes:
            combo.set('')

    def create_sheet(self, notebook, existing_data):
        # Main frame
        frame = Frame(notebook, bg="white")  # Consider a background color for clarity
        frame.pack(fill="both", expand=True, pady=20)
        notebook.add(frame, text=existing_data[0])

        # Central frame
        center_frame = Frame(frame, bg="white")
        center_frame.pack(pady=20)

        # Function to create a section (used thrice for existing samples, current samples, and reference gene)
        def create_section(parent, title_text, col):
            label = Label(parent, text=title_text, font=('Helvetica', 10, 'bold'), bg="white")
            label.grid(column=col, row=0, pady=(10, 0))
            section_frame = Frame(parent, bg="white")
            section_frame.grid(column=col, row=1, padx=20, pady=10, sticky=W + E + N + S)
            return section_frame

        # Existing Sample Section
        sample_frame = create_section(center_frame, "Existing Sample Order", 0)

        # Displaying samples
        for num in range(1, len(existing_data[1])):
            for i in range(2):
                text = existing_data[1][0] if i == 0 else existing_data[1][num]
                sample_name = Label(sample_frame, text=text, bg="white")
                sample_name.grid(column=i, row=num, padx=5, pady=5, sticky=W + E)
                self.created_widgets.append(sample_name)

        # Current Sample Order Section
        new_sample_frame = create_section(center_frame, "Current Sample Order", 1)

        # Dropdowns for Current Sample Order
        fc_targets = self.data[1]
        comboboxes_in_sheet = []
        self.combinations.append([existing_data[0], []])
        for num in range(1, len(existing_data[1])):
            pair = []
            for i in range(2):
                var = StringVar()
                drop = ttk.Combobox(new_sample_frame, textvariable=var, state="readonly", values=fc_targets)
                drop.grid(column=i, row=num, padx=5, pady=5, sticky=W + E)
                pair.append(drop)
                drop.bind("<BackSpace>", self.clear_content)
                self.created_widgets.append(drop)
                comboboxes_in_sheet.append(drop)
            self.combinations[-1][1].append(pair)

        # Reference Gene Section
        ref_sample_frame = create_section(center_frame, "Reference Gene", 2)

        ref_var = StringVar()
        ref_drop = ttk.Combobox(ref_sample_frame, textvariable=ref_var, state="readonly", values=self.data[2])
        ref_drop.grid(column=0, row=0, padx=5, pady=5, sticky=W + E)
        ref_drop.bind("<BackSpace>", self.clear_content)
        comboboxes_in_sheet.append(ref_drop)
        self.combinations[-1].append(ref_drop)

        # Clear All Button
        clear_button = ttk.Button(frame, text="Clear All",
                                  command=lambda: self.clear_all_comboboxes(comboboxes_in_sheet))
        clear_button.pack(side="bottom", anchor="se", pady=10, padx=10)

        # Widgets list
        self.created_widgets.extend(
            [frame, sample_frame, new_sample_frame, ref_sample_frame, ref_drop, clear_button])

        # Set default value for reference gene based on existing data
        default_num = None
        for num in range(len(self.data[2])):
            if self.data[2][num] == existing_data[0].split(" - ")[1]:
                default_num = num
                break
        else:
            messagebox.showerror("Error", "Invalid data layout")
            raise reg_qpcr.InvalidDataLayoutException

        return ref_drop, default_num

    def add_combo_button(self):
        if self.mode.get() in ["qPCR ΔΔCт", "qPCR ΔΔCт - Continuous"]:
            label_text = "Samples"
            control_label_text = "Control Group"
            treatment_label_text = "Experimental Group"
        elif self.mode.get() == "ChIP qPCR":
            label_text = "Graph Targets"
            control_label_text = "Control Group"
            treatment_label_text = "Experimental Group"
        else:
            label_text = "Error"
            control_label_text = "Error"
            treatment_label_text = "Error"

        label = Label(self.inner_frame, text=label_text, font=('Helvetica', 12, 'bold'))
        label.grid(column=0, row=self.current_row(1), pady=(20, 0))
        self.created_widgets.append(label)

        if self.mode.get() == "qPCR ΔΔCт - Continuous":
            self.combo_container = Frame(self.inner_frame, bg="white")
            self.combo_container.grid(row=self.current_row(1), column=0, sticky=W + E + N + S, padx=10, pady=10)
            self.created_widgets.append(self.combo_container)

            if self.existing_filename.get() == "Select Existing File":
                self.show_error("No Existing File Selected")
                raise Exception("No Existing File Selected")
            try:
                existing_data = reg_qpcr.get_existing_info(self.existing_filename.get())
            except InvalidFileException:
                self.show_error("Wrong File Type")
                raise Exception("Wrong File Type")

            notebook = ttk.Notebook(self.combo_container)
            notebook.pack(expand=True, fill=BOTH, padx=0, pady=0)
            for num in range(len(existing_data)):
                ref_drop_list = self.create_sheet(notebook, existing_data[num])
                if ref_drop_list[1] is not None:
                    ref_drop_list[0].current(ref_drop_list[1])
            self.created_widgets.append(notebook)
        else:
            self.combo_container = Frame(self.inner_frame, bg="white", bd=2, relief=SUNKEN, padx=10, pady=10)
            self.combo_container.grid(row=self.current_row(1), column=0, sticky=W + E + N + S, padx=10, pady=10)
            self.created_widgets.append(self.combo_container)

            center_frame = Frame(self.combo_container, bg="white", padx=5, pady=5)
            center_frame.pack(side="top", anchor="center")

            self.combo_frame = Frame(center_frame, bg="lightgrey")
            self.combo_frame.grid(column=0, row=self.current_row(0))

            label_frame = Frame(self.combo_frame, padx=5, pady=5)
            label_frame.grid(column=0, row=self.current_row(), sticky=W + E)
            control_label = Label(label_frame, text=control_label_text)
            control_label.grid(column=0, row=self.current_row(), padx=(0, 60))
            treatment_label = Label(label_frame, text=treatment_label_text)
            treatment_label.grid(column=2, row=self.current_row())
            button = Button(center_frame, text="Add More Samples", width=25,
                            command=lambda: self.add_combos(self.data))
            button.grid(column=1, row=self.current_row(0), sticky=W + E + N + S)

            self.created_widgets.extend([center_frame, self.combo_frame, button])

            self.add_combos(self.data)

    def create_focus_handlers(self, placeholder_text):
        def on_focus_in(event):
            entry_widget = event.widget
            if entry_widget.get() == placeholder_text:
                entry_widget.delete(0, "end")
                entry_widget.config(fg="black")

        def on_focus_out(event):
            entry_widget = event.widget
            if not entry_widget.get():
                entry_widget.insert(0, placeholder_text)
                entry_widget.config(fg="gray")

        return on_focus_in, on_focus_out

    def on_window_click(self, event):
        # If the clicked widget is not in the list of entry widgets, remove focus from all entry widgets
        if event.widget not in self.entry_widgets:
            self.focus_set()

    def generate_file_area(self):
        self.bottom_frame = Frame(self.inner_frame)
        self.bottom_frame.grid(row=self.current_row(), column=0)
        self.output_filename = os.path.basename(self.filename.get()).split(".")[0]

        email_button = Button(self.bottom_frame, text="Email Spreadsheet", width=30,
                              command=lambda: self.email_spreadsheet())
        email_button.grid(column=1, row=self.current_row(), pady=10, sticky=W)

        if self.mode.get() == "qPCR ΔΔCт - Continuous":
            download_button_text = "Save To Existing Spreadsheet"
        else:
            download_button_text = "Save New Spreadsheet"

        download_button = Button(self.bottom_frame, text=download_button_text, width=30,
                                 command=lambda: self.save_existing_excel_file())
        download_button.grid(column=0, row=self.current_row(), pady=10, sticky=E)

        self.entry_widgets.extend([self.bottom_frame, email_button, download_button])
        self.created_widgets.extend([self.bottom_frame, email_button, download_button])

    def setup_entry_area(self):
        self.email_mainframe = Frame(self.inner_frame)
        self.email_mainframe.grid(row=self.current_row(0), column=0, pady=20)
        self.email_mainframe.columnconfigure(0, minsize=300)
        self.entry_frame = Frame(self.email_mainframe, padx=10, pady=10)
        self.entry_frame.grid(row=self.current_row(), column=1, sticky=W + E)
        self.entry_label = Label(self.entry_frame, text="Enter Email:")
        self.entry_label.grid(row=self.current_row(), column=1)
        self.entry_combobox = ttk.Combobox(self.entry_frame, values=self.load_previous_entries())
        self.entry_combobox.grid(row=self.current_row(), column=1, sticky=W + E)
        self.entry_widgets.append(self.entry_combobox)
        self.entry_combobox.bind('<Return>', lambda event: self.add_label())
        self.add_button = Button(self.entry_frame, text="Add Email", command=self.add_label)
        self.add_button.grid(row=self.current_row(0), column=3)

        self.created_widgets.extend(
            [self.entry_frame, self.email_mainframe, self.entry_label, self.entry_combobox, self.add_button])

    def setup_labels_area(self):
        self.labels_container = Frame(self.email_mainframe, bg="white", bd=2, relief=SUNKEN, padx=10, pady=10,
                                      width=250)
        self.labels_container.grid(row=self.current_row(1), column=0, sticky=W + E + N + S, padx=10, pady=10)
        self.labels_frame = Frame(self.labels_container, bg="lightgrey")
        self.labels_frame.grid(sticky=W + E + N + S)
        self.created_widgets.append(self.labels_container)
        self.created_widgets.append(self.labels_frame)

    def load_previous_entries(self):
        try:
            with open("previous_entries.txt", "r") as file:
                temp = file.read().strip().split("\n")[:5]
                temp.reverse()
                return temp

        except FileNotFoundError:
            return []

    def save_entry(self, text):
        entries = self.load_previous_entries()
        if text not in entries:
            entries.insert(0, text)
        entries = entries[:5]
        entries.reverse()
        with open("previous_entries.txt", "w") as file:
            file.write("\n".join(entries))

    def create_label_with_close_button(self, text):
        if text not in self.existing_labels:
            self.existing_labels.add(text)
            frame = Frame(self.labels_frame, bg="lightgrey", padx=5, pady=5)
            label = Label(frame, text=text, bg="lightgrey", font=("Courier", 10))
            label.grid(row=self.current_row(), column=0, sticky=W)
            close_button = Button(frame, text="✖", command=lambda: self.destroy_label(frame, text), bg="lightgrey",
                                  fg="red", relief=FLAT, font=("Courier", 10))
            close_button.grid(row=self.current_row(), column=1)
            frame.grid(sticky=W + E)
            self.emails.append(text)

    def destroy_label(self, frame, text):
        self.existing_labels.remove(text)
        self.emails.remove(text)
        frame.destroy()

    def add_label(self):
        text = self.entry_combobox.get()
        if text:
            self.save_entry(text)
            self.create_label_with_close_button(text)
            self.entry_combobox.delete(0, END)
            self.entry_combobox['values'] = self.load_previous_entries()

    def select_existing_file(self):
        file_label = Label(self.choose_file_frame, textvariable=self.existing_filename)
        file_label.grid(column=0, row=0, sticky=W + E)

        upload_button = Button(self.choose_file_frame, text="Choose Existing File", command=lambda: (self.handle_upload_cont()), width=16)
        upload_button.grid(column=1, row=0, padx=50, sticky="W")

        self.created_widgets.extend([file_label, upload_button])

    def handle_upload_cont(self):
        self.current_row_value = 3
        self.clear_widgets()
        self.select_existing_file()
        open_file(self.existing_filename)

        if self.filename.get() == "Select qPCR Results File" or self.existing_filename.get() == "Select Existing File":
            pass
        else:
            self.emails = []
            self.hkgs = []
            self.combinations = []
            if self.mode.get() == "ChIP qPCR":
                self.make_orientation_dropdown()
            try:
                self.acquire_data(self.filename.get())
            except reg_qpcr.InvalidExcelLayoutException as e:
                messagebox.showerror("Error", str(e))
                raise e
            # self.setup_targets_area()
            # self.get_targets(self.data)
            self.add_combo_button()

            self.setup_entry_area()
            self.setup_labels_area()
            self.generate_file_area()
            self.mode_selection.bind('<<ComboboxSelected>>', lambda _: self.switch_modes())
            self.file_uploaded = True
            self.cont_file_uploaded = True

    def handle_upload(self):
        self.current_row_value = 3
        self.clear_widgets()
        open_file(self.filename)
        if self.mode.get() == "qPCR ΔΔCт - Continuous":
            self.select_existing_file()

        if self.filename.get() == "Select qPCR Results File":
            pass
        elif self.mode.get() == "qPCR ΔΔCт - Continuous":
            if self.filename.get() == "Select qPCR Results File" or self.existing_filename.get() == "Select Existing File":
                pass
            else:
                self.emails = []
                self.hkgs = []
                self.combinations = []
                if self.mode.get() == "ChIP qPCR":
                    self.make_orientation_dropdown()
                    self.current_row_value += 1
                try:
                    self.acquire_data(self.filename.get())
                except reg_qpcr.InvalidExcelLayoutException as e:
                    messagebox.showerror("Error", str(e))
                    raise e
                if self.mode.get() != "qPCR ΔΔCт - Continuous":
                    self.setup_targets_area()
                    self.get_targets(self.data)
                self.add_combo_button()

                self.setup_entry_area()
                self.setup_labels_area()
                self.generate_file_area()
                self.mode_selection.bind('<<ComboboxSelected>>', lambda _: self.switch_modes())
                self.file_uploaded = True
        else:
            self.hkgs = []
            self.combinations = []
            if self.mode.get() == "ChIP qPCR":
                self.make_orientation_dropdown()
                self.current_row_value += 1
            try:
                self.acquire_data(self.filename.get())
            except reg_qpcr.InvalidExcelLayoutException as e:
                messagebox.showerror("Error", str(e))
                raise e
            if self.mode.get() != "qPCR ΔΔCт - Continuous":
                self.setup_targets_area()
                self.get_targets(self.data)
            self.add_combo_button()

            self.setup_entry_area()
            self.setup_labels_area()
            self.generate_file_area()
            self.mode_selection.bind('<<ComboboxSelected>>', lambda _: self.switch_modes())
            self.file_uploaded = True

    def switch_modes(self):
        self.current_row_value = 2
        self.clear_widgets()

        if self.mode.get() == "qPCR ΔΔCт - Continuous":
            self.select_existing_file()

        if self.file_uploaded:
            if self.mode.get() == "ChIP qPCR":
                self.make_orientation_dropdown()
                self.current_row_value += 1

            self.emails = []
            self.hkgs = []
            self.combinations = []
            try:
                self.acquire_data(self.filename.get())
            except reg_qpcr.InvalidExcelLayoutException as e:
                messagebox.showerror("Error", str(e))
                raise e
            if self.mode.get() != "qPCR ΔΔCт - Continuous":
                self.setup_targets_area()
                self.get_targets(self.data)
            self.add_combo_button()

            self.setup_entry_area()
            self.setup_labels_area()
            self.generate_file_area()

    def switch_orientation(self):
        self.current_row_value = 3
        self.clear_widgets([self.orientation_dropdown, self.orientation_frame, self.orientation_label])
        self.current_row_value += 1
        self.emails = []
        self.hkgs = []
        self.combinations = []
        try:
            self.acquire_data(self.filename.get())
        except reg_qpcr.InvalidExcelLayoutException as e:
            messagebox.showerror("Error", str(e))
            raise e
        self.setup_targets_area()
        self.get_targets(self.data)
        self.add_combo_button()

        self.setup_entry_area()
        self.setup_labels_area()
        self.generate_file_area()
        self.created_widgets.extend([self.orientation_dropdown, self.orientation_frame, self.orientation_label])

    def make_orientation_dropdown(self):
        self.orientation_frame = Frame(self.inner_frame)
        self.orientation_frame.grid(column=0, row=self.current_row(0))
        self.orientation_label = Label(self.orientation_frame, text="Sample Orientation",
                                       font=('Helvetica', 12, 'bold'))
        self.orientation_label.grid(column=0, row=0, pady=(20, 5))
        self.orientation = StringVar()
        self.orientation_dropdown = ttk.Combobox(self.orientation_frame, textvariable=self.orientation,
                                                 values=["Horizontal", "Vertical"],
                                                 state="readonly")
        self.orientation_dropdown.current(0)
        self.orientation_dropdown.grid(column=0, row=1, pady=(0, 20))
        self.created_widgets.extend([self.orientation_frame, self.orientation_dropdown, self.orientation_label])

        self.orientation_dropdown.bind('<<ComboboxSelected>>', lambda event: self.switch_orientation())

    def email_spreadsheet(self):
        if self.mode.get() == "qPCR ΔΔCт - Continuous":
            try:
                reg_qpcr.write_wb_cont(data=self.data[0],
                                       fold_change_targets=self.get_combinations(),
                                       output_filename=self.existing_filename.get())
            except reg_qpcr.InvalidDataLayoutException as e:
                messagebox.showerror("Error", str(e))
                raise e
            except Exception:
                self.show_error("Unexpected error occurred, check your sample order")
                raise Exception
            self.send_emails(self.existing_filename.get())
            messagebox.showinfo("File Emailed", "Your Excel File has been emailed")
        else:
            # check for exceptions here
            if ".xlsx" in self.output_filename:
                file = self.output_filename
            elif ".xls" in self.output_filename:
                file = self.output_filename.split(".xls")[0] + ".xlsx"
            else:
                file = self.output_filename + ".xlsx"
            try:
                if self.mode.get() == "qPCR ΔΔCт":
                    if len(self.get_hkgs()) == 0:
                        self.show_error("No HKGs Chosen")
                        raise Exception("No HKGs Chosen")
                    reg_qpcr.write_wb(self.data[0], self.get_hkgs(), self.get_combinations(), file)
                elif self.mode.get() == "ChIP qPCR":
                    if len(self.get_hkgs()) == 0:
                        self.show_error("No References Chosen")
                        raise Exception("No References Chosen")
                    chip_qpcr.write_wb(data=self.data[0], reference_targets=self.get_hkgs(),
                                       graph_targets=self.get_combinations(), output_filename=file)
            except IndexError:
                self.show_error("No Combinations Selected or Wrong Mode Selected")
                raise Exception("No Combinations Selected or Wrong Mode Selected")
            self.send_emails(file)
            os.remove(file)
            messagebox.showinfo("File Emailed", "Your Excel File has been emailed")

    def get_hkgs(self):
        hkg_vals = []
        for i in self.hkgs:
            if i[0].get() == 1:
                hkg_vals.append(i[1])
        return hkg_vals

    def apply_to_nested(self, nested_list, func):
        if isinstance(nested_list, list):
            return [self.apply_to_nested(item, func) for item in nested_list]
        else:
            if isinstance(nested_list, str):
                return nested_list
            else:
                return func(nested_list)

    def get_combinations(self):
        true_combos = self.apply_to_nested(self.combinations, lambda x: x.get())
        return true_combos

    def send_emails(self, file):
        file_list = []
        if ".xlsx" in file:
            file = file
        elif ".xls" in file:
            file = file.split(".xls")[0] + ".xlsx"
        else:
            file = file + ".xlsx"
        file_list.append(file)
        file_list.append(self.filename.get())
        sender_email = SENDER_EMAIL
        app_password = APP_PASSWORD
        send_to = ""
        cc = ''
        for i in self.emails:
            send_to += f",{i}"
        if send_to == "":
            self.show_error("No Emails Entered")
            raise Exception("No Emails Entered")
        send_email.email_excel(files=file_list, password=app_password, send_from=sender_email, send_to=send_to, cc=cc)

    def save_existing_excel_file(self):
        if self.mode.get() == "qPCR ΔΔCт - Continuous":
            try:
                reg_qpcr.write_wb_cont(data=self.data[0],
                                       fold_change_targets=self.get_combinations(),
                                       output_filename=self.existing_filename.get())
                messagebox.showinfo("File Updated", "Your Excel File has been updated")
            except reg_qpcr.InvalidDataLayoutException as e:
                messagebox.showerror("Error", str(e))
                raise e
            except Exception:
                self.show_error("Unexpected error occurred, check your sample order")
                raise Exception

        else:
            if ".xlsx" in self.output_filename:
                file = self.output_filename
            elif ".xls" in self.output_filename:
                file = self.output_filename.split(".xls")[0] + ".xlsx"
            else:
                file = self.output_filename + ".xlsx"

            try:
                if self.mode.get() == "qPCR ΔΔCт":
                    if len(self.get_hkgs()) == 0:
                        self.show_error("No HKGs Chosen")
                        raise Exception("No HKGs Chosen")
                    reg_qpcr.write_wb(self.data[0], self.get_hkgs(), self.get_combinations(), file, first_time=True)
                elif self.mode.get() == "ChIP qPCR":
                    if len(self.get_hkgs()) == 0:
                        self.show_error("No References Chosen")
                        raise Exception("No References Chosen")
                    chip_qpcr.write_wb(data=self.data[0], reference_targets=self.get_hkgs(),
                                       graph_targets=self.get_combinations(), output_filename=file)
            except IndexError:
                self.show_error("No Combinations Selected or Empty Combinations Present")
                raise Exception("No Combinations Selected or Empty Combinations Present")

            if not file or not file.endswith('.xlsx'):
                self.show_error("Invalid Excel file path")
                return

            # Open the save file dialog
            filename = filedialog.asksaveasfilename(
                initialdir="/",
                initialfile=self.output_filename,
                defaultextension=".xlsx",
                filetypes=[("Excel Workbook", ".xlsx"), ("All Files", "*.*")],
            )

            # Check if the user canceled the save dialog
            if not filename:
                os.remove(file)
                return

            try:
                # Copy the original file to the chosen location
                shutil.copy(file, filename)
            except Exception as e:
                self.show_error(f"An error occurred while saving the file: {e}")

            os.remove(file)


@log_exceptions
def run_app():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    run_app()
