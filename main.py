# HT-BDS: Hub-Temperature-Controlled BDS Testing Software
from random import randint
from dataclasses import dataclass
import time
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox # add file dialog and message box imports
import openpyxl
from typing import Literal
import matplotlib.pyplot as plt
from matplotlib.axes import Axes
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk) # type: ignore
import os
import numpy as np
import pandas as pd
from enum import IntFlag, auto
import threading
try:
    import devices
except Exception as e:
        print(f"Error importing devices module: {e}")

# make a folder for output files if it doesn't already exist
RUNNING_PATH = os.path.abspath(os.getcwd())
OUTPUT_FOLDER = "HT-BDS"
OUTPUT_FILEPATH = os.path.join(RUNNING_PATH, OUTPUT_FOLDER)
os.makedirs(OUTPUT_FILEPATH, exist_ok=True)

DEVICE_NAMES = [dev.name for dev in devices.DEVICE_TYPE_LIST]

class RUN_STATE(IntFlag): # to intflag since it allows for binary combinations of states
    IDLE            = auto() # App just begun, nothing asked yet. Accept user inputs, parameters, etc.
    PROGRAMMING     = auto() # Inputs now become outputs. Begin instructing connected devices on what signals/commands to expect.
    READY           = auto() # After machines have been programmed successfully, review details before comitting to run.
    PAUSE           = auto() # Paused run. User has requested the run be held in place for intervention or analysis.
    TEMP_CHANGING   = auto() # Running. Temperatures are changing towards the set temperature.
    PROBE_SWITCHING = auto() # Running. Probes are switching for switching frequency measurements.
    LCR_MEASURING   = auto() # Running. LCR is sweeping frequncies for a given probe.
    DONE            = auto() # Run complete / End of programs
    RUNNING         = TEMP_CHANGING | PROBE_SWITCHING | LCR_MEASURING # Combination of states that would be considered 'running'

# dataclass for storing the run configuration, which can be easily passed around and modified as needed
@dataclass(frozen=True)
class RunConfig:
    device : devices.Device | None
    start_temp: float
    step_temp: float
    max_temp: float
    dwell_time: float
    heat_rate: float
    focus_freq: float

# Initialize global constants
CHAR_OHM    = '\u03A9'
CHAR_THETA  = '\u0398'
CHAR_DEG    = '\u00b0'
CHAR_DEGC   = '\u00b0C'
CHAR_MU     = '\u03bc'

UNITS: dict[str, float] = {
    'f':        1e-15,
    'p':        1e-12,
    'n':        1e-9,
    CHAR_MU:    1e-6,
    'u':        1e-6,
    'm':        1e-3,
    'c':        1e-2,
    '':         1,
    'k':        1e+3,
    'M':        1e+6,
    'G':        1e+9,
    'T':        1e+12,
}

# column names
#TEMP_STEP_COLUMNS=("Step #", "Time [min]", f"Set Temp. [{CHAR_DEGC}]")
FREQ_STEP_COLUMNS=("Step #", "Frequency [Hz]", "*")
TEST_DATA_COLUMNS=("Probe #", f"Temp. [{CHAR_DEGC}]","Freq. [Hz]","Cp [F]","Df [1]", f"ESR [{CHAR_OHM}]") #add esr
TEMPERATURE_READINGS_COLUMNS=('Time [s]','Cycle','Mode',f'Set Temp [{CHAR_DEGC}]',f'Chamber Temp [{CHAR_DEGC}]',f'User Temp [{CHAR_DEGC}]')
TEMP_PLAN_COLUMNS=("Step #", "Target Temp. [°C]", "Dwell Time [min]", "Ramp Time [min]", "Elapsed Time [min]")

# takes a python list and converts it into a string and reconverts it back to a comma separated list
class ListVar(tk.StringVar): # pars
    type: str
    name: str
    # value:  list
    # var:    tk.StringVar
    def __init__(self, type: Literal['int', 'bool', 'float', 'str']='float', name:str=""): #, value=[]):
        self.type = type
        self.name = name
        # self.value = value
        super().__init__(name=name)
        # self.var = tk.StringVar()
        # self.var.trace_add("read", self.get)
    
    # coverts list to a string
    def set(self, value:list[int|float|str|bool]=[]): # type: ignore
        print(f"Setting {self.name} var... {value}")
        super().set(','.join([str(entry) for entry in value]))
        # self.value = value
    
    # takes the string, splits it by comma, and converts it back to a list of the correct type
    def get(self) -> list[int|float|str|bool]: # type: ignore
        try:
            # super().get()
            # Remove spaces and split by comma
            # return super().get()
            values = super().get().split(',')
            print(f"Getting {self.name} var... {values}")
            raw_list = [entry.strip() for entry in values if entry.strip()]
            match (self.type):
                case 'int':
                    return [int(int_entry) for int_entry in raw_list]
                case 'float':
                    return [float(float_entry) for float_entry in raw_list]
                case 'str':
                    return [str_entry for str_entry in raw_list]
                case 'bool':
                    return [bool(bool_entry) for bool_entry in raw_list]
        except ValueError:
            return []  # Invalid input
        
# highlights all text when entry is selected, and formats the text when deselected
class Entry(tk.Entry):
    textvariable: tk.Variable
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if 'textvariable' in kwargs.keys():
            self.textvariable = kwargs['textvariable']
        self.bind("<FocusIn>", self.focus_highlight)
        self.bind("<FocusOut>", self.format_input)
        
    def focus_highlight(self, *args):
        self.selection_range(0, 'end')
        
    def format_input(self, *args):
        self.textvariable.set(self.textvariable.get())
# alternating color for table rows, with function to update the whole table based on a given dataframe
class Table(ttk.Treeview):
    def __init__(self, header_widths, increment=True, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.tag_configure('evenrow', background='#E8E8E8')
        self.tag_configure('oddrow', background='#FFFFFF')
        self.index_inc = increment
        columns = self.cget('columns')
        self.set_headings(columns, header_widths)
    
    # sets the column widths and headings based on the given lists
    def set_headings(self, column_list, width_list):
        self.column('#0', width=0, stretch=True)
        self.heading('#0', text="", anchor='w')
        for i, col in enumerate(column_list):
            width = width_list[i] if i < len(width_list) else 80
            self.column(col, minwidth=50, width=width, stretch=False, anchor='w')
            self.heading(col, text=col, anchor='w')
    
    # updates the whole table based on the given dataframe, with alternating row colors
    def update_table(self, dataframe: pd.DataFrame):
        # Add data with alternating row colors
        # TODO: Fix reiterating whole table for everything ADDED APPEND ROW 
        self.delete(*self.get_children())
        for i, data in enumerate(dataframe.itertuples(index=False, name=None)):
            if self.index_inc:
                inc_data = [data[0]+1]
                inc_data.extend(data[1:])
                data = inc_data
            if i % 2 == 0:
                self.insert(parent='', index='end', values=data, tags='evenrow')
            else:
                self.insert(parent='', index='end', values=data, tags='oddrow')

    # adds a single row to the end of the table, with alternating row color
    def appendRow(self, data):
        i = len(self.get_children())
        if self.index_inc:
            data = [i+1] + list(data)
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        self.insert(parent='', index='end', values=data, tags=tag)
        print("Appending row to table:", data)

# makes the temperature step plot, updates it, and draws it on the canvas when called
class TempStepPlot:
    master:             tk.Tk
    plot_figure:        Figure
    plot_canvas:        FigureCanvasTkAgg
    temperature_axes:   Axes
    times:              ListVar
    temperature_data:   ListVar
    
    def __init__(self, master_window):
        self.master = master_window
        self.plot_figure = Figure(figsize=(5, 4), dpi=80, layout='constrained')
        self.plot_canvas = FigureCanvasTkAgg(self.plot_figure, master_window)
        self.plot_figure.patch.set_facecolor('#F0F0F0')
        self.temperature_axes = self.plot_figure.subplots(1,1)
        self.times = ListVar('float', "Times")
        self.temperature_data = ListVar('float', "Temperatures")
    
    # remade to update the plot based on a given dataframe of the temperature plan, with proper formatting and labels
    def update_plot(self, dataframe: pd.DataFrame, start_temp: float = 25.0, *args):
        print("Updating Temperature Plot...")
        ax = self.temperature_axes
        ax.clear()
        ax.set_title("Temperature Plan Preview")
        ax.set_xlabel("Elapsed Time [min]")
        ax.set_ylabel(r"$T_{set}$" + f" [{CHAR_DEGC}]")
        ax.grid(which='both')

        if dataframe.empty:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
            self.plot_canvas.draw()
            return

        times = [0.0]
        temps = [start_temp]
        current_time=0.0
        for _, row in dataframe.iterrows():
            target = float(row[TEMP_PLAN_COLUMNS[1]])
            dwell_time = float(row[TEMP_PLAN_COLUMNS[2]])
            ramp_time = float(row[TEMP_PLAN_COLUMNS[3]])

            ramp_end = current_time + ramp_time
            dwell_end = ramp_end + dwell_time

            times.append(ramp_end)
            temps.append(target)

            times.append(dwell_end)
            temps.append(target)
            current_time = dwell_end

        if times[-1] > 120:
            times = [t / 60 for t in times]
            ax.set_xlabel("Elapsed Time [hr]")

        ax.plot(times, temps, marker='o', linestyle='-')
        self.plot_canvas.draw()

# remade this to include esr plots and handle general vs frequency and vs temperature plots
class TestDataPlot:
        def __init__(self, master_window):
            self.master = master_window
            self.notebook = ttk.Notebook(master_window)
            self.widget = self.notebook

            self.figures ={}
            self.canvases = {}
            self.axes = {}

            self.build_tab(
                "frequency",
                "Versus Frequency",
                [ 
                    ("Capacitance vs Frequency", 'Frequency [Hz]', 'Cp [F]'),
                    ("Dissipation vs Frequency", 'Frequency [Hz]', 'Df [1]'),
                    ("ESR vs Frequency", 'Frequency [Hz]', f"ESR [{CHAR_OHM}]"),
                ],
            )
            self.build_tab(
                "temperature",
                "Versus Temperature",
                [ 
                    ("Capacitance vs Temperature", 'Temperature [°C]', 'Cp [F]'),
                    ("Dissipation vs Temperature", 'Temperature [°C]', 'Df [1]'),
                    ("ESR vs Temperature", 'Temperature [°C]', f"ESR [{CHAR_OHM}]"),
                ],
            )

        def build_tab(self, tab_key, tab_title, plot_specs):
            frame = tk.Frame(self.notebook)
            self.notebook.add(frame, text=tab_title)
            fig = Figure(figsize = (6,7), dpi = 80, layout="constrained")
            fig.patch.set_facecolor('#F0F0F0')
            axes = fig.subplots(3,1)

            canvas = FigureCanvasTkAgg(fig, frame)
            canvas.get_tk_widget().pack(fill='both', expand=True)

            toolbar = NavigationToolbar2Tk(canvas, frame)
            toolbar.update()
            toolbar.pack(side='bottom', fill='x')

            for ax, (plot_title, x_label, y_label) in zip(axes, plot_specs):
                self._format_empty_axis(ax, plot_title, x_label, y_label)

            # Store references to the figure, canvas, and axes for later updates
            self.figures[tab_key] = fig
            self.canvases[tab_key] = canvas
            self.axes[tab_key] = axes

            canvas.draw()
        
        def update_plots(self, dataframe = None):
            if dataframe is None or dataframe.empty:
                self._draw_empty()
                return
            
            # check if all required columns are present in the dataframe, if not, draw empty and return
            required_columns = list(TEST_DATA_COLUMNS)
            for col in required_columns:
                if col not in dataframe.columns:
                    print(f"Dataframe is missing required column: {col}")
                    self._draw_empty()
                    return

            # frequency plots
            freq_col = "Freq. [Hz]"
            temp_col = f"Temp. [{CHAR_DEGC}]"
            cp_col = "Cp [F]"
            df_col = "Df [1]"
            esr_col = f"ESR [{CHAR_OHM}]"
            
            freq_axes = self.axes["frequency"]
            freq_specs = [
                (freq_axes[0], dataframe[freq_col], dataframe[cp_col], "Capacitance vs Frequency", "Frequency [Hz]", "Cp [F]", "semilogx"),
                (freq_axes[1], dataframe[freq_col], dataframe[df_col], "Dissipation Factor vs Frequency", "Frequency [Hz]", "Df [1]", "semilogx"),
                (freq_axes[2], dataframe[freq_col], dataframe[esr_col], "ESR vs Frequency", "Frequency [Hz]", f"ESR [{CHAR_OHM}]", "semilogx"),
            ]

            for ax, x, y, title, xlabel, ylabel, mode in freq_specs:
                ax.clear()
                ax.set_title(title)
                ax.set_xlabel(xlabel)
                ax.set_ylabel(ylabel)
                ax.grid(which='both')
                if mode == "semilogx":
                    ax.semilogx(x, y, marker='o', linestyle='-')
                else:
                    ax.plot(x, y, marker='o', linestyle='-')
            
            self.canvases["frequency"].draw()

            # temperature plots
            temp_axes = self.axes["temperature"]
            temp_specs = [
                (temp_axes[0], dataframe[temp_col], dataframe[cp_col], "Capacitance vs Temperature", f"Temperature [{CHAR_DEGC}]", "Cp [F]"),
                (temp_axes[1], dataframe[temp_col], dataframe[df_col], "Dissipation Factor vs Temperature", f"Temperature [{CHAR_DEGC}]", "Df [1]"),
                (temp_axes[2], dataframe[temp_col], dataframe[esr_col], "ESR vs Temperature", f"Temperature [{CHAR_DEGC}]", f"ESR [{CHAR_OHM}]"),
            ]
        
            for ax, x, y, title, xlabel, ylabel in temp_specs:
                ax.clear()
                ax.set_title(title)
                ax.set_xlabel(xlabel)
                ax.set_ylabel(ylabel)
                ax.grid(which="both")
                ax.plot(x, y, marker="o", linestyle="-")

            self.canvases["temperature"].draw()

        # if no data or missing columns, draw empty plot
        def _draw_empty(self):
            for tab_key, axes in self.axes.items():
                for ax in axes:
                    title = ax.get_title()
                    xlabel = ax.get_xlabel()
                    ylabel = ax.get_ylabel()
                    self._format_empty_axis(ax, title, xlabel, ylabel)
                self.canvases[tab_key].draw()

        # helper function to format an axis as empty with a "No Data" message
        def _format_empty_axis(self, ax, title, xlabel, ylabel):
            ax.clear()
            ax.set_title(title)
            ax.set_xlabel(xlabel)
            ax.set_ylabel(ylabel)
            ax.grid(which='both')
            ax.text(
                0.5,
                0.5,
                "No Data",
                ha='center',
                va='center',
                transform=ax.transAxes
            )

# Main application class to hold all the state and functions for the app
class App():
    app_root:                   tk.Tk
    state:                      RUN_STATE
    # new variables
    max_temp_dvar:              tk.DoubleVar # for max temperature in stepwise temp cycle
    step_temp_dvar:             tk.DoubleVar # for step temperature in stepwise temp cycle
    dwell_time_dvar:            tk.DoubleVar # for dwell time at each step in stepwise temp cycle
    heat_rate_dvar:             tk.DoubleVar # for heat ramp rate in stepwise temp cycle

    # temperature step variables
    temp_step_locked:           bool
    temp_step_lock_list:        list[tk.Button | Entry]
    temp_step_table:            Table
    temp_step_data:             pd.DataFrame
    temp_step_plot:             TempStepPlot

    # Frequency sweep variables
    first_freq_dvar:            tk.DoubleVar
    last_freq_dvar:             tk.DoubleVar
    custom_freq_dvar:           tk.DoubleVar
    points_per_decade_ivar:     tk.IntVar
    freq_step_locked:           bool
    freq_step_lock_list:        list[tk.Button | Entry]
    freq_step_table:            Table
    freq_step_data:             pd.DataFrame
    custom_freq_table:          Table
    custom_freq_data:           pd.DataFrame

    # test management variables
    active_probes_listvar:      ListVar
    active_lines_listvar:       ListVar
    test_data:                  pd.DataFrame

    device_list:                list[devices.Device]
    active_device:              devices.Device
    device_strvar:              tk.StringVar
    message_strvar:             tk.StringVar
    response_strvar:            tk.StringVar

    temperature_readings_table:     pd.DataFrame
    temperature_rolling_table:      pd.DataFrame
    reading_lock:               threading.Lock
    rolling_lock:               threading.Lock
    reading_lock:               threading.Lock
    last_temperature_readings_length:   int

    def __init__(self):
        # Init app vars/entries/etc.
        self.app_root = tk.Tk()
        self.state = RUN_STATE.DONE

        # new variables
        self.start_temp = tk.DoubleVar(value=25.0)
        self.step_temp = tk.DoubleVar(value=20.0)
        self.max_temp = tk.DoubleVar(value=125.0)
        self.dwell_time = tk.DoubleVar(value=5.0)
        self.heat_rate = tk.DoubleVar(value=2.5)

        self.first_freq_dvar = tk.DoubleVar(value=devices.LCR_MIN_FREQ)
        self.last_freq_dvar = tk.DoubleVar(value=devices.LCR_MAX_FREQ)
        self.points_per_decade_ivar = tk.IntVar(value=10)
        self.custom_freq_dvar = tk.DoubleVar()
        
        self.temp_step_data = pd.DataFrame(columns=TEMP_PLAN_COLUMNS)
        self.custom_freq_data = pd.DataFrame(columns=FREQ_STEP_COLUMNS[1:])
        self.freq_step_data = pd.DataFrame(columns=FREQ_STEP_COLUMNS[1:])
        self.test_data = pd.DataFrame(columns=TEST_DATA_COLUMNS)
        
        self.device_strvar = tk.StringVar(value=devices.DEVICE_TYPE_LIST[0].name)
        self.message_strvar = tk.StringVar()
        self.response_strvar = tk.StringVar()
        self.device_list = [None]*len(devices.DEVICE_TYPE_LIST) # type: ignore
        for index, dev in enumerate(devices.DEVICE_TYPE_LIST):
            try:
                self.device_list[index] = dev()
                print(f"{DEVICE_NAMES[index]} successfully connected!")
            except:
                self.device_list[index] = None # type: ignore
                print(f"{DEVICE_NAMES[index]} could not connect!")
        self.active_device = self.device_list[0]
        
        self.temperature_readings_table = pd.DataFrame(columns=TEMPERATURE_READINGS_COLUMNS) # master table; huge after hours
        self.temperature_rolling_table = pd.DataFrame(columns=TEMPERATURE_READINGS_COLUMNS) # rolling table, ~30 entries
        self.readings_lock = threading.Lock()
        self.rolling_lock = threading.Lock()
        self.last_temperature_readings_length = 0
        
        self.build_app()
        self.app_root.mainloop()
        
    # switched to a paned window setup
    def build_app(self):
        # Do some setup
        self.app_root.title("Temperature-Controlled BDS Testing")
        # self.app_window.iconphoto(True, tk.PhotoImage(file="peak-nano-logo-blue.png"))
        #self.app_root.iconbitmap("peak-nano-logo-blue.ico")
        self.app_root.state('zoomed')
        self.app_root.protocol("WM_DELETE_WINDOW", self.on_closing)
        plt.ioff()
        self.build_menubar(self.app_root)
        outer = tk.PanedWindow(
        self.app_root,
        orient="horizontal",
        sashrelief="raised",
        sashwidth=6,
        bg="#1E1E2E"
        )
        outer.pack(fill="both", expand=True, padx=6, pady=6)

        left = tk.LabelFrame(outer, text="Test Setup", font=('default', 14, 'bold'), padx=10, pady=10)
        middle = tk.LabelFrame(outer, text="Test Management", font=('default', 14, 'bold'), padx=10, pady=10)
        right = tk.LabelFrame(outer, text="Data Plots", font=('default', 14, 'bold'), padx=10, pady=10)

        outer.add(left, minsize=300, width=360)
        outer.add(middle, minsize=340, width=420)
        outer.add(right, minsize=400, width=900)

        self.build_test_setup(left)
        self.build_test_management(middle)
        self.build_data_plots(right)
        return self.app_root
    # END build_app()

    # function for confirmation before closing
    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.app_root.destroy()
    
    def padding(self, master, x=0, y=0, side='top', fill='both', expand=False, **kwargs):
        tk.Frame(master, width=x, height=y, **kwargs).pack(side=side, fill=fill, expand=expand) # type: ignore
    # fix for updated vars and removal of old temp cycle inputs
    def sync_temp_step(self, dataframe: pd.DataFrame, *args):
        self.temp_step_table.update_table(dataframe)
        self.temp_step_plot.update_plot(dataframe)
    # menubar with file options
    def build_menubar(self, master):
        menubar = tk.Menu(master, tearoff=False)
        fileMenu = tk.Menu(menubar, tearoff=False)
        fileMenu.add_command(label="Export Results to Excel") #, command=self.export_results) TODO: Add export function
        fileMenu.add_separator()
        fileMenu.add_command(label="Exit", command=self.on_closing)
        menubar.add_cascade(label="File", menu=fileMenu)
        self.app_root.config(menu=menubar)
    
    def build_test_setup(self, master):
        notebook_style = ttk.Style()
        # notebook_style.theme_create(themename="NotebookStyle", parent='default', settings={
        #     "TNotebook.Tab": {'configure': {'font' : ('default', '12', 'normal')},}})
        # notebook_style.theme_use("NotebookStyle")
        notebook = ttk.Notebook(
            master,
            style='TNotebook',
        ); notebook.pack(side='top', fill='both', expand=True)

        # Temperature Tab
        temperature_tab = tk.Frame(
            master=notebook, 
            width=400, 
            height=280, 
            # background='green',
        ); temperature_tab.pack(side='top', fill='both', expand=True)
        notebook.add(temperature_tab, text="Temperature")
        self.build_test_temp_tab(temperature_tab)
        
        # Frequency Tab 
        frequency_tab = tk.Frame(
            master=notebook, 
            width=400, 
            height=280,
        ); frequency_tab.pack(side='top', fill='both', expand=True)
        notebook.add(frequency_tab, text="Frequency")
        self.build_test_freq_tab(frequency_tab)
        
        # (Device Communication) Testing Tab
        devices_tab = tk.Frame(
            master=notebook, 
            width=400, 
            height=280,
        ); devices_tab.pack(side='top', fill='both', expand=True)
        notebook.add(devices_tab, text="Devices") #renamed
        self.build_devices_tab(devices_tab)
        
        # (Probe Selection) Probe Tab 
        probes_tab = tk.Frame(
            master=notebook, 
            width=400, 
            height=280,
        ); probes_tab.pack(side='top', fill='both', expand=True)
        notebook.add(probes_tab, text="Probes")
        tk.Label(master=probes_tab, text="Probe Selection", font=('default', 16, 'bold')).pack(pady=(20,10))
        tk.Label(master=probes_tab, text="Implementation coming soon!", justify="center", wraplength=380).pack() # TODO: Add probe selection functions
        self.build_probes_tab(probes_tab)

    # moving generate temperature plan into app
    def generate_temperature_plan(self):
        try:
            start_temp = self.start_temp.get()
            step_temp = self.step_temp.get()
            max_temp = self.max_temp.get()
            dwell_time = self.dwell_time.get()
            heat_rate = self.heat_rate.get()

            rows = []

            current_temp = start_temp
            elapsed = 0
            step = 1
            target = start_temp + step_temp

            while target <= max_temp:
                ramp_time = abs(target - current_temp) / heat_rate
                elapsed += ramp_time + dwell_time

                rows.append([
                    step,
                    target,
                    dwell_time,
                    ramp_time,
                    elapsed
                ])

                current_temp = target
                target += step_temp
                step += 1

            self.temp_step_data = pd.DataFrame(rows, columns=TEMP_PLAN_COLUMNS)

            self.temp_step_table.update_table(self.temp_step_data)
            self.temp_step_plot.update_plot(self.temp_step_data)

        except ValueError as e:
            messagebox.showerror("Invalid Temperature Plan", str(e))

    # helper function to create labeled entries
    def _labeled_entry(self, parent, label: str, variable: tk.Variable, width=10):
            row = tk.Frame(parent)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, width=22, anchor="w").pack(side="left")
            entry = Entry(row, textvariable=variable, width=width, justify="right")
            entry.pack(side="left")
            return entry
    
    # remake
    def build_test_temp_tab(self, temperature_tab):
        params = tk.LabelFrame(temperature_tab, text="Step Parameters", padx=8, pady=8)
        params.pack(fill="x", pady=6)

        self._labeled_entry(params, f"Start Temp [{CHAR_DEGC}]", self.start_temp)
        self._labeled_entry(params, f"Step Size [{CHAR_DEGC}]", self.step_temp)
        self._labeled_entry(params, f"Max Temp [{CHAR_DEGC}]", self.max_temp)
        self._labeled_entry(params, "Dwell Time [min]", self.dwell_time)
        self._labeled_entry(params, f"Heat Ramp [{CHAR_DEGC}/min]", self.heat_rate)

        tk.Button(
            params,
            text="Generate Plan",
            command=self.generate_temperature_plan,
            fg="royalblue"
        ).pack(pady=(6, 0))

        table_frame = tk.LabelFrame(temperature_tab, text="Temperature Plan", padx=6, pady=6)
        table_frame.pack(fill="both", expand=True, pady=4)

        table_with_scroll_frame = tk.Frame(table_frame)
        table_with_scroll_frame.pack(fill="both", expand=True)

        self.temp_step_table = Table(
            master=table_with_scroll_frame,
            columns=TEMP_PLAN_COLUMNS,
            displaycolumns='#all',
            show='headings',
            selectmode='none',
            height=8,
            header_widths=[40, 80, 80, 80, 80, 80],
            increment=False,
        )
        self.temp_step_table.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(
            master=table_with_scroll_frame,
            orient='vertical',
            command=self.temp_step_table.yview,
        )
        scrollbar.pack(side='left', fill='y')
        self.temp_step_table.configure(yscrollcommand=scrollbar.set)

        plot_frame = tk.LabelFrame(temperature_tab, text="Plan Preview", padx=6, pady=6)
        plot_frame.pack(fill="both", expand=True, pady=4)

        self.temp_step_plot = TempStepPlot(plot_frame)
        self.temp_step_plot.plot_canvas.get_tk_widget().pack(fill='both', expand=True)
        self.temp_step_plot.update_plot(self.temp_step_data)

    def build_test_freq_tab(self, frequency_tab):
        sweep_params_box = tk.LabelFrame(
            master=frequency_tab,
            text="Primary Frequencies: Sweep",
            font=('default', 10),
            padx=10,
            pady=10,
        )
        sweep_params_box.pack(side='top', fill='x')
        frequency_logspace_box = tk.Frame(
            master=sweep_params_box,
        )
        frequency_logspace_box.pack(side='top', fill='x')
        
        low_freq_frame = tk.Frame(
            master=frequency_logspace_box,
        )
        low_freq_frame.pack(side='left')
        low_freq_label = tk.Label(
            master=low_freq_frame,
            text="First [Hz] "
        )
        low_freq_label.pack(side='left', fill='y')
        low_freq_entry = Entry(
            master=low_freq_frame,
            width=10,
            justify='right',
            textvariable=self.first_freq_dvar,
        )
        low_freq_entry.pack(side='left', fill='y')
        
        high_freq_frame = tk.Frame(
            master=frequency_logspace_box,
        )
        high_freq_frame.pack(side='left')
        high_freq_label = tk.Label(
            master=high_freq_frame,
            text=" to Last [Hz] "
        )
        high_freq_label.pack(side='left', fill='y')
        high_freq_entry = Entry(
            master=high_freq_frame,
            width=10,
            justify='right',
            textvariable=self.last_freq_dvar,
        )
        high_freq_entry.pack(side='left', fill='y')
        
        points_per_decade_frame = tk.Frame(
            master=frequency_logspace_box,
        )
        points_per_decade_frame.pack(side='left')
        points_per_decade_label = tk.Label(
            master=points_per_decade_frame,
            text=" @ Points/Dec. "
        )
        points_per_decade_label.pack(side='left', fill='y')
        points_per_decade_entry = Entry(
            master=points_per_decade_frame,
            width=7,
            justify='right',
            textvariable=self.points_per_decade_ivar,
        )
        points_per_decade_entry.pack(side='left', fill='y')
        # points_per_decade_entry.bind("<Return>", lambda *args: print(self.first_freq_dvar.get(), self.last_freq_dvar.get(), self.points_per_decade_ivar.get()))
        
        def sync_freq_tables():
            custom_df = self.custom_freq_data
            logspace_df = self.freq_step_data
            self.custom_freq_table.update_table(self.custom_freq_data)
            # custom_df[FREQ_STEP_COLUMNS[2]] = '*'
            # logspace_df[FREQ_STEP_COLUMNS[2]] = ''
            df = pd.concat([logspace_df, custom_df], ignore_index=True)
            df = df.sort_values(by=FREQ_STEP_COLUMNS[1])
            df = df.reset_index(drop=True)
            self.freq_step_table.update_table(df)
        
        def set_freq_logspace_pressed():
            minimum_frequency = self.first_freq_dvar.get()
            maximum_frequency = self.last_freq_dvar.get()
            points_per_decade = self.points_per_decade_ivar.get()
            print(minimum_frequency, maximum_frequency, points_per_decade)
            low_decade = np.log10(minimum_frequency)
            high_decade = np.log10(maximum_frequency)
            num_decades = high_decade - low_decade
            if num_decades < 1:
                data_points = np.linspace(minimum_frequency, maximum_frequency, num=points_per_decade)
            else:
                num_points = int(num_decades) * points_per_decade
                data_points = np.logspace(start=low_decade, stop=high_decade, num=num_points, endpoint=True, base=10)
            data_points = [float(f"{x:.6g}") for x in data_points]
            print("Logspace frequencies:", data_points)
            df = self.freq_step_data
            if len(df):
                df = pd.DataFrame(columns=FREQ_STEP_COLUMNS[1:])
            df[FREQ_STEP_COLUMNS[1]] = data_points
            df[FREQ_STEP_COLUMNS[2]] = ''
            self.freq_step_data = df
            self.freq_step_table.update_table(df)
            sync_freq_tables()
            # var.frequency_array = list(data_points)
        
        tk.Frame(master=sweep_params_box, height=10).pack(side='top', fill='x')
        
        logspace_button_box = tk.Frame(
            master=sweep_params_box,
        )
        logspace_button_box.pack(side='top', fill='x')
        
        set_logspace_button = tk.Button(
            master=logspace_button_box,
            text="Set Logspace",
            command=set_freq_logspace_pressed,
        )
        set_logspace_button.pack(side='top')
        points_per_decade_entry.bind("<Return>", lambda *args: set_logspace_button.invoke())

        # Frequency list tables
        tk.Frame(master=frequency_tab, height=10).pack(side='top', fill='x')
        both_freq_table_frame = tk.Frame(
            master=frequency_tab,
        )
        both_freq_table_frame.pack(side='top', fill='both', expand=True)
        left_side_box = tk.LabelFrame(
            master=both_freq_table_frame,
            text="Secondary Frequencies",
            font=('default', 10),
            padx=10,
            pady=10,
        )
        left_side_box.pack(side='left', fill='y', expand=True)
        tk.Frame(master=both_freq_table_frame, width=10).pack(side='left')
        right_side_box = tk.LabelFrame(
            master=both_freq_table_frame,
            text="Frequency Steps",
            font=('default', 10),
            padx=10,
            pady=10,
        )
        right_side_box.pack(side='left', fill='y', expand=True)
        
        # Test settings button inputs
        setting_buttons_box = tk.Frame(
            master=left_side_box,
        )
        setting_buttons_box.pack(side='top', fill='x')
        
        manual_freq_frame = tk.Frame(
            master=setting_buttons_box,
        )
        manual_freq_frame.pack(side='top', fill='x')
        manual_freq_label = tk.Label(
            master=manual_freq_frame,
            text="Manual [Hz] "
        )
        manual_freq_label.pack(side='left')
        manual_freq_entry = Entry(
            master=manual_freq_frame,
            width=10,
            justify='right',
            textvariable=self.custom_freq_dvar,
        )
        manual_freq_entry.pack(side='left')
        button_row = tk.Frame(
            master=setting_buttons_box,
            padx=10,
            pady=10,
        )
        button_row.pack(side='top', fill='x', expand=True)
        
        def add_manual_freq_pressed(*args):
            print("add step")
            freq = self.custom_freq_dvar.get()
            df = self.custom_freq_data
            if (freq in self.freq_step_data[FREQ_STEP_COLUMNS[1]]):
                print(f"{freq} Hz exists in sequence!")
                return
            print([freq, '*'])
            print(self.custom_freq_data)
            self.custom_freq_data.loc[len(df)] = [freq, '*']
            print(self.custom_freq_data)
            sync_freq_tables()
            
        add_setting_button = tk.Button(
            master=button_row,
            text="Add Step",
            command=add_manual_freq_pressed,
        )
        add_setting_button.pack(side='left', expand=True)
        manual_freq_entry.bind("<Return>", lambda *args: add_setting_button.invoke())
        
        def drop_manual_freq_pressed(*args):
            df = self.custom_freq_data
            if len(df) == 0:
                return
            self.custom_freq_data = df.drop(df.tail(1).index)
            sync_freq_tables()
        
        remove_setting_button = tk.Button(
            master=button_row,
            text="Drop Step",
            command=drop_manual_freq_pressed,
        )
        remove_setting_button.pack(side='left', expand=True)
        manual_freq_entry.bind("<Shift-Return>", lambda *args: remove_setting_button.invoke())
        
        def clear_manual_freqs_pressed(*args):
            df = self.custom_freq_data
            if len(df) == 0:
                return
            self.custom_freq_data = pd.DataFrame(columns=FREQ_STEP_COLUMNS[1:])
            sync_freq_tables()
        
        clear_setting_button = tk.Button(
            master=setting_buttons_box,
            text="Clear List",
            command=clear_manual_freqs_pressed,
        )
        clear_setting_button.pack(side='top', expand=True)
        
        # Custom freq table
        tk.Frame(master=left_side_box, height=10).pack(side='top', fill='x')
        
        custom_table_labelframe = tk.Frame(
            master=left_side_box,
        )
        custom_table_labelframe.pack(side='top', fill='y', expand=True)
        custom_table_with_scroll_frame = tk.Frame(
            master=custom_table_labelframe,
        )
        custom_table_with_scroll_frame.pack(side='left', fill='y', expand=True)
        self.custom_freq_table = Table(
            master=custom_table_with_scroll_frame,
            columns=FREQ_STEP_COLUMNS[:2],
            displaycolumns='#all',
            show='headings',
            selectmode='none',
            height=20,
            header_widths=[40,100],
            increment=True,
        )
        self.custom_freq_table.pack(side='left', fill='both', expand=True)
        custom_table_scrollbar = ttk.Scrollbar(
            master=custom_table_with_scroll_frame, 
            orient='vertical', 
            command=self.custom_freq_table.yview,
        )
        custom_table_scrollbar.pack(side='left', fill='y')
        self.custom_freq_table.configure(yscrollcommand=custom_table_scrollbar.set)
        
        # Full list table
        full_table_labelframe = tk.Frame(
            master=right_side_box,
        )
        full_table_labelframe.pack(side='left', fill='y', expand=True)
        full_table_with_scroll_frame = tk.Frame(
            master=full_table_labelframe,
        )
        full_table_with_scroll_frame.pack(side='left', fill='y', expand=True)
        self.freq_step_table = Table(
            master=full_table_with_scroll_frame,
            columns=FREQ_STEP_COLUMNS,
            displaycolumns='#all',
            show='headings',
            selectmode='none',
            height=20,
            header_widths=[40,100,16],
            increment=True,
        )
        self.freq_step_table.pack(side='left', fill='both', expand=True)
        full_table_scrollbar = ttk.Scrollbar(
            master=full_table_with_scroll_frame, 
            orient='vertical', 
            command=self.freq_step_table.yview, 
        )
        full_table_scrollbar.pack(side='left', fill='y')
        self.freq_step_table.configure(yscrollcommand=full_table_scrollbar.set)

    def build_devices_tab(self, devices_tab):
        self.padding(devices_tab, y=10, side='top')
        device_labelframe = tk.LabelFrame(
            master=devices_tab,
            text="Active Devices",
            font=('default', 12),
            padx=10,
            pady=10,
        ); device_labelframe.pack(side='top', fill='x')
        
        def set_device(*args):
            index = DEVICE_NAMES.index(self.device_strvar.get())
            self.active_device = self.device_list[index]
            
        device_combobox = ttk.Combobox(
            master=device_labelframe,
            textvariable=self.device_strvar,
            values=DEVICE_NAMES,
        ); device_combobox.pack(side='left', fill='x', expand=True)
        self.device_strvar.trace_add('write', set_device)
        
        self.padding(devices_tab, y=10, side='top')
        message_labelframe = tk.LabelFrame(
            master=devices_tab,
            text="Manual Command",
            font=('default', 12),
            padx=10,
            pady=10,
        ); message_labelframe.pack(side='top', fill='x')
        # message_label = tk.Label(
        #     master=message_labelframe,
        #     text="CMD: ",
        # ); message_label.pack(side='left')
        message_entry = Entry(
            master=message_labelframe,
            justify='left',
            textvariable=self.message_strvar,
        ); message_entry.pack(side='left', fill='x', expand=True)
        
        try:
            # self.device_list = devices.SunSystemsOven_EC1A()
            pass
        except:
            pass # TODO: Separate process looping device connection/communication
        
        self.padding(message_labelframe, x=10, side='left')
        send_button = tk.Button(
            master=message_labelframe,
            text="Send",
            command=lambda *args: threading.Thread(target=self.device_msg).start(),
        ); send_button.pack(side='right')
        message_entry.bind("<Return>", lambda *args:send_button.invoke())
        
        self.padding(devices_tab, y=10, side='top')
        response_frame = tk.LabelFrame(
            master=devices_tab,
            text="Response",
            font=('default', 12),
            padx=10,
            pady=10,
        ); response_frame.pack(side='top', fill='both')
        response_box = tk.Text(
            master=response_frame,
            width=10,
        ); response_box.pack(side='top', fill='both')
        self.response_strvar.trace_add('write', lambda *args: response_box.replace("1.0", 'end', self.response_strvar.get()))
    # END build_test_setup()
    
    def build_probes_tab(self, probes_tab):
        # TODO: Implement this tab for BDS functionality
        self.padding(probes_tab, y=10, side='top')
        probes_labelframe = tk.LabelFrame(
            master=probes_tab,
            text="Probes",
            font=('default', 12),
            padx=10,
            pady=10,
        ); probes_labelframe.pack(side='top', fill='x')
    
    def device_msg(self, query:str="", expected:str="", device:devices.Device|None=None, hushed:bool=False):
        if device == None:
            device = self.active_device
            if device == None:
                if not hushed: print(f"Device not set! Cannot send: {query}!")
                return (-1, np.nan)
        if query:
            command = query
        else:
            command = self.message_strvar.get()
        (code, reply) = device.send(cmd=command, expect=expected)
        display_text = f"Length: {str(code)} -->\n{reply}"
        self.response_strvar.set(display_text)
        if not hushed: print(f"Command: '{command}' ({code})\n\tReturn: '{reply}'")
        return (code, reply)

    def build_data_plots(self, master):
        
        self.test_plot = TestDataPlot(master)
        self.test_plot.widget.pack(side="top", fill="both", expand=True)
        self.test_plot.update_plots(self.test_data)
    # END build_data_plots()

    def build_test_management(self, master):
        controls_labelframe = tk.LabelFrame(
            master,
            text = "Controls",
            font=('default', 12),
            padx=10,
            pady=10,
        ) 
        controls_labelframe.pack(side='top', fill='x')
        button_box = tk.Frame(
            master=controls_labelframe,
            padx=10,
            pady=10,
            bg="#E8E8E8",
        ); button_box.pack(side='top', fill='both')
        
        self.padding(button_box, x=10, side='left', bg='#E8E8E8')
        
        program_button = tk.Button(
            master=button_box,
            text="Program",
            command=lambda *args: threading.Thread(target=lambda: self.program_device(self.get_run_data()), daemon=True).start(),
            font=('default',10,'bold'),
            foreground='royalblue',
            width=10,
        ); program_button.pack(side='left')
        
        self.padding(button_box, x=20, side='left', bg='#E8E8E8')
        
        def run_button_pressed():
            if run_button.config('relief')[-1] == 'sunken': # If button pressed
                run_button.config(relief='raised', bg='SystemButtonFace', fg='forestgreen', text="Run")
            else:
                run_button.config(relief='sunken', bg='forestgreen', fg='white', text="Running")
                thread = threading.Thread(target=self.run(self.get_run_data()))
                thread.start()
                program_button.config(state='disabled', relief='flat')
                run_button.config(state='disabled', relief='flat')
                # pause_button.config(state='disabled', relief='flat')
                thread.join()
        run_button = tk.Button(
            master=button_box,
            text="Run",
            command=lambda *args: threading.Thread(target=run_button_pressed, daemon=True).start(),
            font=('default',10,'bold'),
            foreground='forestgreen',
            width=10,
            # state='disabled',
            # relief='flat',
        ); run_button.pack(side='left')
        
        self.padding(button_box, x=20, side='left', bg='#E8E8E8')
        
        def pause_button_pressed(): # TODO: May need to modify functionality... Oven may trigger it's own breakpoint during a pause, if currently holding near it's set temp
            if pause_button.config('relief')[-1] == 'sunken': # If button pressed
                pause_button.config(relief='raised', bg='SystemButtonFace', fg='darkviolet', text="Pause")
                self.state = RUN_STATE.TEMP_CHANGING
                self.device_msg(device=self.active_device, query='BKPNTC') # not threaded; needs to be directly under main thread control for safety
            else:
                pause_button.config(relief='sunken', bg='purple2', fg='white', text="Resume")
                self.state = RUN_STATE.PAUSE
                self.device_msg(device=self.active_device, query='BKPNT') # not threaded; needs to be directly under main thread control for safety
        pause_button = tk.Button(
            master=button_box,
            text="Pause",
            command=pause_button_pressed,
            font=('default',10,'bold'),
            foreground='darkviolet',
            width=10,
            # state='disabled',
            # relief='flat',
        ); pause_button.pack(side='left')
        
        self.padding(button_box, x=20, side='left', bg='#E8E8E8')
        
        def stop_button_pressed():
            self.state = RUN_STATE.DONE
            self.device_msg(device=self.active_device, query='STOP') # not threaded; needs to be directly under main thread control for safety
            self.device_msg(device=self.active_device, query='COFF') # not threaded; needs to be directly under main thread control for safety
            self.device_msg(device=self.active_device, query='HOFF') # not threaded; needs to be directly under main thread control for safety
        stop_button = tk.Button(
            master=button_box,
            text="Stop",
            command=stop_button_pressed,
            font=('default',10,'bold'),
            foreground='crimson',
            width=10,
        ); stop_button.pack(side='left')
        
        self.padding(button_box, x=10, side='left', bg='#E8E8E8')
        # stop_button.config(state='disabled', relief='flat')
        
        self.padding(master, y=10, side='top')
        
        
        table_labelframe = tk.LabelFrame(
            master,
            text="Data Table",
            font=('default', 12),
            padx=10,
            pady=10,
        ); table_labelframe.pack(side='top', fill='both')
        table_with_scroll_frame = tk.Frame(
            master=table_labelframe,
        ); table_with_scroll_frame.pack(side='top', fill='y')
        data_table = Table(
            master=table_with_scroll_frame,
            columns=TEST_DATA_COLUMNS,
            displaycolumns='#all',
            show='headings',
            selectmode='none',
            height=16,
            header_widths=[60, 60, 60, 60, 60, 60],
        ); data_table.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(table_with_scroll_frame, orient='vertical', command=data_table.yview)
        data_table.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='left', fill='both', expand=True)
        self.padding(master, y=10, side='top')
        # Configure alternating row colors
        # data_table.tag_configure('evenrow', background='#E8E8E8')
        # data_table.tag_configure('oddrow', background='#FFFFFF')
        data = [
            ('a', 'b', 'c'),
            ('A', 'B', 'C'),
            ('aa', 'bb', 'cc'),
            ('aA', 'bB', 'cC'),
            ('Aa', 'Bb', 'Cc'),
            ('AA', 'BB', 'CC'),
            ('aaa', 'bbb', 'ccc'),
            ('aaA', 'bbB', 'ccC'),
            ('aAa', 'bBb', 'cCc'),
            ('aAA', 'bBB', 'cCC'),
            ('Aaa', 'Bbb', 'Ccc'),
            ('AaA', 'BbB', 'CcC'),
            ('AAa', 'BBb', 'CCc'),
            ('AAA', 'BBB', 'CCC'),
        ]

        # Add data with alternating row colors
        for i in range(len(data)):
            if i % 2 == 0:
                data_table.insert(parent='', index=i, values=data[i], tags=('evenrow',))
            else:
                data_table.insert(parent='', index=i, values=data[i], tags=('oddrow',))
        
                
        temp_chart_labelframe = tk.LabelFrame(
            master,
            text="Temperature Chart",
            font=('default', 12),
            padx=10,
            pady=10,
        ); temp_chart_labelframe.pack(side='top', fill='both', expand=True)
        
        temp_cycle_fig = Figure(figsize=(5, 4), dpi=80, layout='constrained')
        temp_cycle_fig.patch.set_facecolor('#F0F0F0')
        self.temp_cycle_axis = temp_cycle_fig.add_subplot()
        axis = self.temp_cycle_axis
        temp_cycle_canvas = FigureCanvasTkAgg(temp_cycle_fig, temp_chart_labelframe)
        temp_cycle_canvas.get_tk_widget().pack(side='top', fill='both', expand=True)
        MAX_LENGTH = 100
        HEADER_LIST = [TEMPERATURE_READINGS_COLUMNS[x] for x in [0,3,4,5]]
        axis.text(0.5, 0.5, "No Data", ha='center', va='center', transform=axis.transAxes)
        
        def apply_cosmetics(ax:Axes,domain='min',times=[0.0],sets=[0.0],chambs=[0.0],users=[0.0]):
            ax.set_title("Oven Readings (100 Samples)")
            axis.set_xlabel(f"Time [{domain}]")
            ax.set_ylabel(f"Sampled Temp [{CHAR_DEGC}]")
            ax.grid(which='both')
            axis.plot(times, sets, linestyle='-', color='blue')
            axis.plot(times, chambs, linestyle=':', color='orange')
            axis.plot(times, users, linestyle='--', color='red')
            axis.legend(labels=["Set","Chamber","User"])
        apply_cosmetics(axis)
            
        def t_update_temps(*args):
            if self.state == RUN_STATE.TEMP_CHANGING:
                master_length = len(self.temperature_readings_table)
                rolling_length = len(self.temperature_rolling_table)
                total_length = master_length + rolling_length
                if total_length - self.last_temperature_readings_length <= 0:
                    print("Updating Temp Cycle Plot... No Changes")
                    pass
                else:
                    print("Updating Temp Cycle Plot... New Data")
                    if total_length > MAX_LENGTH:
                        indicies = list(np.linspace(0, total_length-1, MAX_LENGTH, True, dtype=int))
                        master_indicies = [index for index in indicies if index < master_length]
                        rolling_indicies = [index - master_length for index in indicies if index >= master_length]
                        master_slice = self.temperature_readings_table.loc[master_indicies, HEADER_LIST]
                        rolling_slice = self.temperature_rolling_table.loc[rolling_indicies, HEADER_LIST]
                    else:
                        master_slice = self.temperature_readings_table[HEADER_LIST]
                        rolling_slice = self.temperature_rolling_table[HEADER_LIST]
                    self.last_temperature_readings_length = total_length
                    combined_table = pd.concat([master_slice,rolling_slice]) # indicies unused for now
                    times = combined_table[TEMPERATURE_READINGS_COLUMNS[0]].tolist()
                    sets = combined_table[TEMPERATURE_READINGS_COLUMNS[3]].tolist()
                    chambs = combined_table[TEMPERATURE_READINGS_COLUMNS[4]].tolist()
                    users = combined_table[TEMPERATURE_READINGS_COLUMNS[5]].tolist()
                    last_time = times[-1]
                    if last_time > 3600: # 3600 s = 1 hr
                        times = [x/3600 for x in times]
                        domain = 'hr'
                    elif 60 < last_time <= 3600: # 60 s = 1 min
                        times = [x/60 for x in times]
                        domain = 'min'
                    else:
                        domain = 's'
                        
                    axis.clear()
                    
                    apply_cosmetics(axis,domain,times,sets,chambs,users)
                    temp_cycle_canvas.draw()
                    # print(times)
                    # print(temps)
            temp_cycle_canvas.get_tk_widget().after(1000, t_update_temps)
        
        # update_thread = threading.Thread(target=t_update_temps, args=[], daemon=True)
        temp_cycle_canvas.get_tk_widget().after(1000, t_update_temps)
            
    # END build_data_table()

    def get_run_data(self, *args):
        device = self.active_device
        # if device != None:
        high_temp = self.high_temp_dvar.get()
        high_time = self.high_time_dvar.get()
        low_temp = self.low_temp_dvar.get()
        low_time = self.low_time_dvar.get()
        cycles = self.temp_cycles_ivar.get()
        heat_rate = self.heat_ramp_dvar.get()
        cool_rate = self.cool_ramp_dvar.get()
        data = (device, high_temp, high_time, low_temp, low_time, cycles, heat_rate, cool_rate)
        return data
    
    def program_device(self, vars:tuple[devices.Device, float, float, float, float, int, float, float]):
        self.state = RUN_STATE.PROGRAMMING
        device: devices.Device = vars[0]
        print("Programming ...")
        # print("Device:", device.name)
        if device is None:
            print("... Programming failed! No device attached!")
        if type(device) == devices.SunSystemsOven_EC1A:
            high_temp = vars[1]
            high_time = vars[2]
            low_temp = vars[3]
            low_time = vars[4]
            cycles = vars[5]
            heat_rate = vars[6]
            cool_rate = vars[7]
            
            # Send program to oven
            self.device_msg(device=device, query="ON")
            self.device_msg(device=device, query="DELP0")
            self.device_msg(device=device, query="STORE0")
            self.device_msg(device=device, query="EDIT0")
            self.device_msg(device=device, query="HON")
            # self.device_msg(device=device, query="CON")
            self.device_msg(device=device, query="SINT=NNNNNNYNNY0")
            self.device_msg(device=device, query=f"FOR I0=1,{cycles+1},+")
            self.device_msg(device=device, query=f"RATE={heat_rate:.2f}")
            self.device_msg(device=device, query=f"WAIT={int(high_time//60):02d}:{int(high_time%60):02d}:{int(high_time*60%60):02d}")
            self.device_msg(device=device, query=f"SET={high_temp:.1f}")
            self.device_msg(device=device, query="BKPNT I0")
            self.device_msg(device=device, query=f"RATE={cool_rate:.2f}")
            self.device_msg(device=device, query=f"WAIT={int(low_time//60):02d}:{int(low_time%60):02d}:{int(low_time*60%60):02d}")
            self.device_msg(device=device, query=f"SET={low_temp:.1f}")
            self.device_msg(device=device, query="BKPNT I0")
            self.device_msg(device=device, query="NEXT I0")
            self.device_msg(device=device, query="HOFF")
            self.device_msg(device=device, query="COFF")
            self.device_msg(device=device, query="END")
            print("... Programming successful!")
            
        self.state = RUN_STATE.READY
    
    def run(self, vars:tuple[devices.Device, float, float, float, float, int, float, float]):
        self.state = RUN_STATE.TEMP_CHANGING
        device: devices.Device = vars[0]
        print("Starting Run ...")
        # print("Device:", device.name)
        if type(device) == devices.SunSystemsOven_EC1A:
            high_temp = vars[1]
            high_time = vars[2]
            low_temp = vars[3]
            low_time = vars[4]
            cycles = vars[5]
            heat_rate = vars[6]
            cool_rate = vars[7]
            
            # Calc time to get from low temp to high temp, shortest wait during first heat
            max_wait = int((high_temp - low_temp) / min(heat_rate, cool_rate)*60 + 1) # duration in seconds; with an extra minute just in case
            
            write_lock = threading.Lock() # uneccessary now, but will be helpful later when viewing data live
            stop_read = threading.Event()
            stop_excel = threading.Event()
            cycle_duration = ((high_temp-low_temp)/heat_rate + high_time + (high_temp-low_temp)/cool_rate + low_time)*60 # seconds of runtime per cycle
            FIDELITY = 0.1 # 10% 
            READ_INTERVAL = 10# FIDELITY*cycle_duration # seconds between data reads from the oven
            SAVE_INTERVAL = 30 # seconds between saves of table data to excel file
            
            def blank_rolling():
                return pd.DataFrame(columns=TEMPERATURE_READINGS_COLUMNS)
            # TODO: Check that old data was removed safely
            self.temperature_rolling_table = blank_rolling()
            
            def t_oven_read(stop_flag:threading.Event, start:float, details:tuple):
                while not stop_flag.is_set():
                    chamber_temp_response = self.device_msg(device=device, query="CHAM?", hushed=True)
                    user_temp_response = self.device_msg(device=device, query="USER?", hushed=True)
                    chamber = chamber_temp_response[1]
                    user = user_temp_response[1]
                    current_time = time.time() - start
                    try:
                       chamber = float(chamber)
                       user = float(user)
                    except:
                        pass
                    # debug = randint(0,99)
                    data = [current_time, details[0], details[1], details[2], chamber, user]
                    with write_lock:
                        pos = max(len(self.temperature_rolling_table), 0) # default to 0 if not started yet
                        self.temperature_rolling_table.loc[pos] = data # append latest data
                    print("Thread @", pos, data)
                    time.sleep(READ_INTERVAL)
            
            def t_write_data(stop_flag:threading.Event, filename:str):
                master_index = 0
                counter = 0
                self.temperature_rolling_table.to_excel(filename, sheet_name="Temperature Data", index=True, header=True) # Get the basic excel file format down, a.k.a. headers are in there
                while not stop_flag.is_set():
                    if counter < SAVE_INTERVAL:
                        print(f"Excel sleeping {counter+1}/{SAVE_INTERVAL}")
                        counter += 1
                        time.sleep(1)
                        continue
                    with write_lock:
                        table = self.temperature_rolling_table
                        entries_to_write = len(table)
                        self.temperature_rolling_table = blank_rolling()
                    if entries_to_write <= 0: # return to waiting if the table is empty
                        print("Excel saw no changes to write")
                    else:
                        print("EXCEL UPDATE:", master_index, "+", entries_to_write)
                        table.index += master_index
                        with pd.ExcelWriter(path=filename,engine='openpyxl',mode='a',if_sheet_exists='overlay') as writer:
                            table.to_excel(writer, sheet_name="Temperature Data", index=True, header=False, startrow=master_index+1, na_rep='nan')
                        self.temperature_readings_table = pd.concat([self.temperature_readings_table, table], ignore_index=True)
                        master_index += entries_to_write
                    counter = 0
                    time.sleep(1)
                # broke out of the loop, test ended; dump remaining data now
                with write_lock:
                    last_table = self.temperature_rolling_table
                    last_entries_to_write = len(last_table)
                last_table.index += master_index
                with pd.ExcelWriter(path=filename,engine='openpyxl',mode='a',if_sheet_exists='overlay') as writer:
                    last_table.to_excel(writer, sheet_name="Temperature Data", index=True, header=False, startrow=master_index+1, na_rep='nan')
                self.temperature_readings_table = pd.concat([self.temperature_readings_table, last_table], ignore_index=True)
                master_index += last_entries_to_write
                self.temperature_rolling_table = blank_rolling()
                print("EXCEL COMPLETE:\n", master_index, "to", master_index + last_entries_to_write)
            
            default_filename = OUTPUT_FILEPATH+f"\\TEST_RUN_{cycles}x{int(high_temp)}-{int(low_temp)}.xlsx"
            output_filename = ""
            while not output_filename:
                output_filename = filedialog.asksaveasfilename(initialdir=OUTPUT_FILEPATH, title="Output File?", defaultextension='.xlsx', filetypes=[('Excel Worksheet', '*.xlsx')])
                print("Output Filename:", output_filename)
                    
            excel_thread = threading.Thread(target=t_write_data, args=[stop_excel, output_filename], daemon=True)
            start_time = time.time()
            
            self.device_msg(device=device, query="RUN 0")
            excel_thread.start()
            for cycle_count in range(1, cycles+1):
                print(f"Heat: {cycle_count}")
                read_thread = threading.Thread(target=t_oven_read, args=[stop_read, start_time, [cycle_count, "Heating", high_temp]], daemon=True)
                read_thread.start()
                device.wait_interrupt(None) # for val in range(0,int(10e9)): pass # device.wait_interrupt(max_wait)
                if self.state == RUN_STATE.PAUSE:
                    device.wait_interrupt(None) # for val in range(0,int(10e9)): pass # device.wait_interrupt(max_wait)
                stop_read.set()
                read_thread.join()
                stop_read.clear()
                self.device_msg(device=device, query="BKPNTC")
                
                print(f"Cool: {cycle_count}")
                read_thread = threading.Thread(target=t_oven_read, args=[stop_read, start_time, [cycle_count, "Cooling", low_temp]], daemon=True)
                read_thread.start()
                device.wait_interrupt(None) # for val in range(0,int(10e9)): pass # device.wait_interrupt(max_wait)
                if self.state == RUN_STATE.PAUSE:
                    device.wait_interrupt(None) # for val in range(0,int(10e9)): pass # device.wait_interrupt(max_wait)
                stop_read.set()
                read_thread.join()
                stop_read.clear()
                self.device_msg(device=device, query="BKPNTC")
                # print("Table:\n", self.temperature_readings_table)
            stop_excel.set()
            excel_thread.join()
            stop_excel.clear()
            # self.temperature_readings_table.to_excel(OUTPUT_FILEPATH+f"TEST_RUN_{cycles}x{int(high_temp)}-{int(low_temp)}_FULL.xlsx", index=True, header=True)
        self.state = RUN_STATE.DONE
        print("... Run Finished!")
        pass
    
    def stop(self):
        self.state = RUN_STATE.DONE
        pass
    

app: App
if __name__ == "__main__":
    print("App started...")
    app = App()
    print("...App ended!")