from tkinter import messagebox, Menu, filedialog
from math import ceil
import customtkinter, os, sys
import json, pathlib, win32api, pyautogui
from PIL import Image
from runtime.data_analysis import DataAnalysis
from runtime.py_constants.months_in_year import months_dictionary, months_abv_dictionary
current_path = str(pathlib.Path(__file__).parent.parent)

customtkinter.set_appearance_mode("Light") 
customtkinter.set_default_color_theme("red")  

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__(icon_index=1)
        version = "v1.0.0"
        self.DataAnalysis = DataAnalysis()

        errors = self.DataAnalysis.errors
    
        if errors:
            messagebox.showerror('Error', errors[0])
            self.destroy()
            exit()

        self.protocol("WM_DELETE_WINDOW", self.save_settings)

        self.settings = self.DataAnalysis.settings

        # Creating Menubar
        menubar = Menu(self)
        
        # Adding File Menu and commands
        file = Menu(menubar, tearoff = 0)
        menubar.add_cascade(label ='File', menu = file)
        file.add_command(label ='Print', command = self.print_file)
        file.add_separator()
        file.add_command(label ='Exit', command = self.destroy)
        
        # Display Menu
        self.config(menu = menubar)

        # Window Title
        self.title("Pierpont Mechanical Business Analytics")

        # Default Window Size
        self.geometry(f"{1450}x{1000}")

        font_manager = customtkinter.FontManager()
        # Load Custom Fonts
        font_manager.windows_load_font(font_path=self.resource_path("assets/fonts/Roboto-Regular.ttf"))
        font_manager.windows_load_font(font_path=self.resource_path("assets/fonts/Roboto-Bold.ttf"))
        font_manager.windows_load_font(font_path=self.resource_path("assets/fonts/AlegreyaSC-Bold.ttf"))
        font_manager.windows_load_font(font_path=self.resource_path("assets/fonts/Serif Gothic Regular.ttf"))

   
        title_font = customtkinter.CTkFont(family='Alegreya SC', size=30, weight="bold")
        address_font = customtkinter.CTkFont(family='Serif Gothic', size=14, weight="normal")
        self.regular_font = customtkinter.CTkFont(family='Roboto', size=14, weight="normal")
        self.large_bold = customtkinter.CTkFont(family='Roboto', size=17, weight="bold")
        self.button_font = customtkinter.CTkFont(family='Roboto', size=13, weight="normal")
        self.bold_font = customtkinter.CTkFont(family='Roboto', size=14, weight="bold")

        # Configure App Grid Layout (4x4)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Intitalize Class Variables
        self.panel_row = 5
        self.panel_rowspan = 1
        self.previous_project_index = 0
        self.page_project = (1, 0)
        self.toplevel_window = None
        self.profit_frame1 = None
        self.profit_frame2 = None
        self.scrollable_labor = None
        self.scrollable_filter = None
        self.scrollable_personel = None

        """ 
            Header Frame 
        """
        self.header_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color=('#700e10','#333333'))
        self.header_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self.header_frame.grid_rowconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=1)
        # Logo Image
        bg_image = customtkinter.CTkImage(Image.open(current_path + "/assets/images/Pierpont_Logo.png"),size=(235, 84))
        bg_image_label = customtkinter.CTkLabel(self.header_frame, text='', image=bg_image)
        bg_image_label.grid(row=0, column=0, padx=20, pady=(10, 10), sticky='nw')
        # Address
        address_label = customtkinter.CTkLabel(self.header_frame, text_color='#f5f5f5', text='250 Fulton Avenue\nGarden City Park, NY 11040\nPhone: (516) 746-7300\nFax: (516) 746-7302', font=address_font, justify='left')
        address_label.grid(row=0, column=3, padx=25, pady=(10, 10), sticky='e')

        """ 
            Main Frame 
        """
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.main_frame.grid(row=1, column=0, rowspan=3, columnspan=4, sticky="nsew")
        self.main_frame.grid_rowconfigure(4, weight=1)
        self.main_frame.grid_columnconfigure((3,5), weight=1)
        # App Title
        title_label = customtkinter.CTkLabel(self.main_frame, text="pierpont mechanical business analytics", fg_color='#222', corner_radius=6, text_color='#f5f5f5', font=title_font)
        title_label.grid(row=0, column=0, columnspan=7, ipadx=10, ipady=1, padx=25, pady=10, sticky='n')

        # Project Navigation
        self.seg_button_1 = customtkinter.CTkSegmentedButton(self.main_frame, font=self.regular_font, command=self.maintenance_construction_navigation)
        self.seg_button_1.grid(row=1, column=5, ipadx=10, padx=(20, 10), pady=(10, 8), sticky="n")
        self.seg_button_1.configure(values=["Maintenance", "Construction"])
        self.seg_button_1.set("Maintenance")

        self.seg_button_2 = customtkinter.CTkSegmentedButton(self.main_frame, font=self.regular_font, command=self.maintenance_construction_navigation)
        self.seg_button_2.grid(row=2, column=5, ipadx=10, padx=(20, 10), pady=(7, 5), sticky="n")
        self.seg_button_2.configure(values=["Project Basis", "Cumulative Basis"])
        self.seg_button_2.set("Project Basis")

        # Settings Button
        settings_image = customtkinter.CTkImage(Image.open(current_path + "/assets/images/settings_icon.png"),size=(16, 16))
        settings_button = customtkinter.CTkButton(self.main_frame, width=35, text="", image=settings_image, command=self.open_toplevel)
        settings_button.grid(row=1, column=0, padx=(30,10), pady=10, sticky='e')

        # Regular Time
        reg_price_label = customtkinter.CTkLabel(self.main_frame, text="Reg ($ / Hr)", corner_radius=5, fg_color=('#ccc','#333333'), font=self.regular_font)
        reg_price_label.grid(row=1, column=1, ipadx=15, padx=(15, 0), pady=(10, 10), sticky='w')

        self.reg_price_entry = customtkinter.CTkEntry(self.main_frame, width=105, font=self.regular_font, border_width=1, border_color='#777')
        self.reg_price_entry.grid(row=1, column=1, padx=(150,10), pady=(10, 10), sticky='w')
        self.reg_price_entry.bind("<Return>", self.recalculate_regular_price)
        self.reg_price_entry.insert(0, '110')
        self.regular_hourly = 110

        # Page Navigation Frame
        self.page_frame = customtkinter.CTkFrame(self.main_frame, corner_radius=0, border_width=1, border_color='#777')
        self.page_frame.grid(row=3, column=5, padx=(10,20), pady=(10, 0), sticky="nsew")
        self.page_frame.grid_rowconfigure(0, weight=1)
        self.page_frame.grid_columnconfigure(1, weight=1)

        self.page_label = customtkinter.CTkLabel(self.page_frame, text="Page 1", font=self.bold_font)
        self.page_label.grid(row=0, column=1, padx=5, pady=10, sticky='nsew')

        self.next_button = customtkinter.CTkButton(self.page_frame, width=80, height=24, text='Next', fg_color='gray40', hover_color='gray25', font=self.button_font, command=self.project_next_button)
        self.next_button.grid(row=0, column=2, padx=(10, 22), pady=10, sticky="e")

        self.previous_button = customtkinter.CTkButton(self.page_frame, width=80, height=24, text='Previous', fg_color='gray40', hover_color='gray25', font=self.button_font, command=self.project_previous_button)
        self.previous_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        # Project Scrollable Frame
        self.scrollable_label_button_frame = ScrollableLabelButtonFrame(master=self.main_frame, command=self.select_project_event, corner_radius=0, border_width=1, border_color='#777')
        self.scrollable_label_button_frame.grid(row=4, column=5, padx=(10,20), rowspan=3, pady=(0, 10), sticky="nsew")

        """ 
            Data Frame 
        """
        self.data_frame = customtkinter.CTkFrame(self.main_frame, corner_radius=6, border_width=1, border_color='#777')
        self.data_frame.grid(row=2, column=0, rowspan=4, columnspan=5, padx=(20,10), pady=(10, 10), sticky="nsew")
        self.data_frame.grid_rowconfigure(5, weight=1)
        self.data_frame.grid_columnconfigure(1, weight=1)

        # Project Info
        self.project_label = customtkinter.CTkLabel(self.data_frame, text="PROJECT: ACOUSTIC META MATERIALS", corner_radius=5, text_color='#f5f5f5', fg_color='#333', font=self.large_bold)
        self.project_label.grid(row=0, column=0, columnspan=3, ipadx=5, ipady=1, padx=(15,10), pady=(10,0), sticky='w')

        self.name_label = customtkinter.CTkLabel(self.data_frame, text="NAME: ACOUSTIC META MATERIALS", font=self.bold_font)
        self.name_label.grid(row=1, column=0, padx=25, pady=(5,0), sticky='w')

        self.address_label = customtkinter.CTkLabel(self.data_frame, text="ADDRESS: 845 THIRD AVENUE FL 18", font=self.bold_font)
        self.address_label.grid(row=1, column=1, padx=15, pady=(5,0), sticky='w')

        self.job_number_label = customtkinter.CTkLabel(self.data_frame, text="JOB #: 001-MS", font=self.bold_font)
        self.job_number_label.grid(row=2, column=0, padx=25, pady=(0,5), sticky='w')

        self.pm_frequency_label = customtkinter.CTkLabel(self.data_frame, text="PM FREQUENCY: MONTHLY", font=self.bold_font)
        self.pm_frequency_label.grid(row=2, column=1, padx=15, pady=(0,5), sticky='w')

        self.service_cost_label = customtkinter.CTkLabel(self.data_frame, text="EMERGENCY BILLABLE ($ / Hr): ", font=self.bold_font)
        self.service_cost_label.grid(row=2, column=1, padx=15, pady=(0,5), sticky='w')

        self.data_seg_button_1 = customtkinter.CTkSegmentedButton(self.data_frame, font=self.regular_font, command=self.job_type_seg_button)
        self.data_seg_button_1.grid(row=3, column=0, columnspan=2, ipadx=15, padx=20, pady=(0, 10), sticky="w")
        self.data_seg_button_1.configure(values=["Maintenance", "Emergency Service"])
        self.data_seg_button_1.set("Maintenance")

        self.data_seg_button_2 = customtkinter.CTkSegmentedButton(self.data_frame, font=self.regular_font, command=self.panel_builder)
        self.data_seg_button_2.grid(row=4, column=0, columnspan=2, ipadx=20, padx=20, pady=20, sticky="w")
        self.data_seg_button_2.configure(values=["Mechanic Name","Filter Cost","Labor Cost","Project Summary"])
        self.data_seg_button_2.set("Project Summary")

        self.data_seg_button_3 = customtkinter.CTkSegmentedButton(self.data_frame, font=self.regular_font, command=self.time_seg_button)
        self.data_seg_button_3.grid(row=3, column=2, ipadx=10, padx=(10,20), pady=(0,10), sticky="")
        self.data_seg_button_3.configure(values=["All Time", "Yearly Basis"])
        self.data_seg_button_3.set("Yearly Basis")

        # Time Frame
        self.time_label_list = []; self.time_button_list = []
        self.time_frame = customtkinter.CTkFrame(master=self.data_frame, corner_radius=0, border_width=1, border_color='#777')
        self.time_frame.grid_rowconfigure(12, weight=1)
        self.time_frame.grid_columnconfigure(1, weight=1)

        # --------------------------------------------------
        # Labor Cost Frame
        self.labor_frame = customtkinter.CTkFrame(self.data_frame, corner_radius=0, border_width=1, border_color='#777')
        self.labor_frame.grid_rowconfigure(1, weight=1)
        self.labor_frame.grid_columnconfigure(0, weight=1)
        
        # --------------------------------------------------
        # Filter Cost Frame
        self.filter_frame = customtkinter.CTkFrame(self.data_frame, corner_radius=0, border_width=1, border_color='#777')
        self.filter_frame.grid_rowconfigure(1, weight=1)
        self.filter_frame.grid_columnconfigure(0, weight=1)
        
        # --------------------------------------------------
        # Profit Frame
        self.profit_frame = customtkinter.CTkFrame(self.data_frame, corner_radius=0, border_width=1, border_color='#777')
        self.profit_frame.grid_rowconfigure(10, weight=1)
        self.profit_frame.grid_columnconfigure(0, weight=1)
       
        # --------------------------------------------------
        # Personel Frame
        self.personel_frame = customtkinter.CTkFrame(self.data_frame, corner_radius=0, border_width=1, border_color='#777')
        self.personel_frame.grid_rowconfigure(1, weight=1)
        self.personel_frame.grid_columnconfigure(0, weight=1)
        
        self.selected_job_type = 'Maintenance'
        self.maintenance_construction_navigation(param="Maintenance")
        self.job_type_seg_button(param='Maintenance')
        self.time_seg_button(param="Yearly Basis")
        
        """ 
            End Section 
        """
        # Refresh Label
        refresh_label = customtkinter.CTkLabel(self.main_frame, text="Refresh", corner_radius=5, fg_color=('#ccc','#333333'), font=self.regular_font)
        refresh_label.grid(row=6, column=4, ipadx=15, padx=(25,80), pady=(10, 20), sticky='')
        # Refresh Button
        refresh_image = customtkinter.CTkImage(Image.open(current_path + "/assets/images/refresh_icon.png"),size=(18, 18))
        refresh_button = customtkinter.CTkButton(self.main_frame, width=40, text="", image=refresh_image, command=self.activate_refresh)
        refresh_button.grid(row=6, column=4, padx=(10,20), pady=(10, 20), sticky='e')
        # UI Scaling
        scaling_label = customtkinter.CTkLabel(self.main_frame, text="Zoom %", corner_radius=5, fg_color=('#ccc','#333333'), font=self.regular_font)
        scaling_label.grid(row=6, column=1, ipadx=15, padx=(10,25), pady=(10, 20), sticky='w')
        scaling_optionemenu = customtkinter.CTkOptionMenu(self.main_frame, width=110, values=["80%", "90%", "100%", "110%", "120%"], fg_color='gray40', button_color='gray40', button_hover_color='gray25', command=self.change_scaling_event, anchor='center', font=self.regular_font, dropdown_font=self.regular_font)
        scaling_optionemenu.grid(row=6, column=1, padx=(120,25), pady=(10, 20), sticky='w')
        scaling_optionemenu.set("100%")
        # App Version
        version_label = customtkinter.CTkLabel(self.main_frame, text=f"{version}", font=customtkinter.CTkFont(family='Roboto', size=12, weight="normal"))
        version_label.grid(row=6, column=0, padx=(20, 0), pady=5, sticky='sw')

    def panel_builder(self, param: str):
        if param == 'Mechanic Name':
            # Personel Frame
            if self.scrollable_personel:
                for widget in self.scrollable_personel.winfo_children(): 
                    widget.destroy()
            for widget in self.personel_frame.winfo_children(): widget.destroy()
            self.personel_frame.grid(row=self.panel_row, column=0, columnspan=2, rowspan=self.panel_rowspan, padx=(20,10), pady=(0, 20), sticky="nsew")
            self.personel_panel_builder()
            self.labor_frame.grid_forget()
            self.profit_frame.grid_forget()
            self.filter_frame.grid_forget()
        elif param == 'Labor Cost':
            # Labor Frame
            if self.scrollable_labor:
                for widget in self.scrollable_labor.winfo_children(): 
                    widget.destroy()
            for widget in self.labor_frame.winfo_children(): widget.destroy()
            self.labor_frame.grid(row=self.panel_row, column=0, columnspan=2, rowspan=self.panel_rowspan, padx=(20,10), pady=(0, 20), sticky="nsew")
            self.labor_panel_builder()
            self.personel_frame.grid_forget()
            self.filter_frame.grid_forget()
            self.profit_frame.grid_forget()
        elif param == 'Filter Cost':
            # Filter Frame
            if self.scrollable_filter:
                for widget in self.scrollable_filter.winfo_children(): 
                    widget.destroy()
            for widget in self.filter_frame.winfo_children(): widget.destroy()
            self.filter_panel_builder()
            self.filter_frame.grid(row=self.panel_row, column=0, columnspan=2, rowspan=self.panel_rowspan, padx=(20,10), pady=(0, 20), sticky="nsew")
            self.labor_frame.grid_forget()
            self.personel_frame.grid_forget()
            self.profit_frame.grid_forget()
        elif param == 'Project Summary':
            # Profit Frame
            if self.profit_frame1:
                for widget in self.profit_frame1.winfo_children(): widget.destroy()
            if self.profit_frame2:
                for widget in self.profit_frame2.winfo_children(): widget.destroy()
            for widget in self.profit_frame.winfo_children(): widget.destroy()
            self.profit_panel_builder()
            self.profit_frame.grid(row=self.panel_row, column=0, columnspan=2, rowspan=self.panel_rowspan, padx=(20,10), pady=(0, 20), sticky="nsew")
            self.filter_frame.grid_forget()
            self.labor_frame.grid_forget()
            self.personel_frame.grid_forget()

    def labor_panel_builder(self):
        mechanic_settings = self.settings['Mechanic Hourly Rates']
        maintenance_or_construction = self.seg_button_1.get()
        project_cumulative = self.seg_button_2.get()

        profit_list = []
        cost_list = []
        billed_cost_list = []
        reg_hours_list = []
        ot_hours_list = []
        dt_hours_list = []
        
        profit_percent_text_length = 0
        profit_text_length = 0
        billed_text_length = 0
        cost_text_length = 0

        if maintenance_or_construction == 'Maintenance' and project_cumulative == 'Project Basis':

            self.scrollable_labor = ScrollablePanelFrame(self.labor_frame, corner_radius=0, border_width=1, border_color='#777')
            self.scrollable_labor.grid(row=1, column=0, columnspan=8, padx=(0,5), pady=(0,0), sticky="nsew")

            visits = sorted(self.work_order_data['Visits'], key=lambda x: x['Date'], reverse=True)

            i = 0
            for visit in visits:
                year = visit['Year']
                month = visit['Month']

                if (self.selected_timeframe == 'All Time' or (self.selected_timeframe == 'Yearly Basis' and self.selected_time == year)) and (visit['Job_Type'] == self.selected_job_type):

                    if visit['Job_Type'] == 'Maintenance':
                        letter = ''; interval = ''
                        if self.maintenance_frequency == 'quarterly':
                            letter = 'Q'
                            interval = ceil(round(int(month) / 3, 1))
                        elif self.maintenance_frequency == 'bi-monthly':
                            letter = 'B'
                            interval = ceil(round(int(month) / 2, 1))
                        elif self.maintenance_frequency == 'monthly':
                            letter = 'M'
                            interval = int(month)
                        elif self.maintenance_frequency == '2x/year':
                            letter = 'T'
                            interval = ceil(round(int(month) / 6, 1))
                        elif self.maintenance_frequency == '3x/year':
                            letter = 'T'
                            interval = ceil(round(int(month) / 4, 1))
                        if letter:
                            maintenance_text = f'{letter}{interval} - '
                        else:
                            maintenance_text = ''
                    else:
                        maintenance_text = ''

                    if self.selected_timeframe == 'All Time':
                        year_text = f" '{year[-2:]}"
                    else: 
                        year_text = ''

                    mechanics = visit['Mechanics']
                    reg = 0
                    ot = 0
                    dt = 0
                    cost = 0
                    for mechanic, hours in mechanics.items():
                        try:
                            hourly_rate = mechanic_settings[mechanic]
                        except:
                            hourly_rate = self.regular_hourly
                        reg += hours['Reg']
                        ot += hours['OT']
                        dt += hours['DT']
                        cost += ((hours['Reg'] * hourly_rate) + (hours['OT'] * hourly_rate * 1.5) + (hours['DT'] * hourly_rate * 2))

                    cost = round(cost)
                    reg_hours_list.append(reg); ot_hours_list.append(ot); dt_hours_list.append(dt); cost_list.append(cost)

                    label = customtkinter.CTkLabel(self.scrollable_labor, text=f"{maintenance_text}{months_dictionary[month]}{year_text}", font=self.regular_font)
                    label.grid(row=i, column=0, padx=25, pady=3, sticky='w')

                    reg_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{reg}", font=self.regular_font)
                    reg_hours.grid(row=i, column=1, padx=25, pady=3, sticky='e')

                    ot_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{ot}", font=self.regular_font)
                    ot_hours.grid(row=i, column=2, padx=25, pady=3, sticky='e')

                    dt_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{dt}", font=self.regular_font)
                    dt_hours.grid(row=i, column=3, padx=25, pady=3, sticky='e')

                    cost_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(cost))))]))
                    cost_comma_str = f'-{cost_comma_str}' if cost < 0 else cost_comma_str
                    cost_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"${cost_comma_str}", font=self.regular_font)
                    cost_hours.grid(row=i, column=4, padx=(25,30), pady=3, sticky='e')    

                    cost_text_length = max(self.regular_font.measure(f'${cost_comma_str}'), cost_text_length)              
                    
                    if self.data_seg_button_1.get() == 'Maintenance':  
                        billed_cost = self.cost_per_visit
                    else:
                        billed_cost = round((reg * self.service_hourly) + (ot * self.service_hourly * 1.5) + (dt * self.service_hourly * 2))

                    billed_cost_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(billed_cost))))]))
                    billed_cost_comma_str = f'-{billed_cost_comma_str}' if billed_cost < 0 else billed_cost_comma_str
                    billed_entry = customtkinter.CTkLabel(self.scrollable_labor, text=f"${billed_cost_comma_str}", font=self.regular_font)
                    billed_entry.grid(row=i, column=5, padx=(25,30), pady=3, sticky='e')  
                    billed_cost_list.append(billed_cost)

                    billed_text_length = max(self.regular_font.measure(f'${billed_cost_comma_str}'), billed_text_length)

                    profit = round(billed_cost_list[-1] - cost)
                    profit_list.append(profit)
                    profit_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(profit))))]))
                    profit_comma_str = f'-{profit_comma_str}' if profit < 0 else profit_comma_str
                    profit_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"${profit_comma_str}", font=self.regular_font)
                    profit_hours.grid(row=i, column=6, padx=(20,20), pady=3, sticky='e')

                    profit_text_length = max(self.regular_font.measure(f"${profit_comma_str}"), profit_text_length)

                    profit_percent = round(profit * 100 / billed_cost_list[-1],1) if billed_cost_list[-1] != 0 else 0
                    text_var = '#137f34' if profit_percent >= 0 else '#ca0b0b'
                    profit_percent_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{profit_percent}%", text_color=text_var, font=self.regular_font)
                    profit_percent_hours.grid(row=i, column=7, padx=(20,15), pady=3, sticky='e')

                    profit_percent_text_length = max(self.regular_font.measure(f"{profit_percent}%"), profit_percent_text_length)

                    i += 1

            if i == 0:
                label = customtkinter.CTkLabel(self.scrollable_labor, text='EMPTY', font=self.regular_font)
                label.grid(row=i, column=0, padx=25, pady=4, sticky='w')

                reg_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0.0", font=self.regular_font)
                reg_hours.grid(row=i, column=1, padx=25, pady=3, sticky='e')

                ot_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0.0", font=self.regular_font)
                ot_hours.grid(row=i, column=2, padx=25, pady=3, sticky='e')

                dt_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0.0", font=self.regular_font)
                dt_hours.grid(row=i, column=3, padx=(25,25), pady=3, sticky='e')

                cost_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"$0", font=self.regular_font)
                cost_hours.grid(row=i, column=4, padx=(25,30), pady=3, sticky='e')  

                billed_entry = customtkinter.CTkLabel(self.scrollable_labor, text=f"$0", font=self.regular_font)
                billed_entry.grid(row=i, column=5, padx=(25,30), pady=3, sticky='e')  

                profit_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"$0", font=self.regular_font)
                profit_hours.grid(row=i, column=6, padx=(20,20), pady=3, sticky='e')

                profit_percent_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0%", font=self.regular_font)
                profit_percent_hours.grid(row=i, column=7, padx=(20,15), pady=3, sticky='e')


            month_label = customtkinter.CTkLabel(self.labor_frame, text=f"Month", font=self.bold_font)
            month_label.grid(row=0, column=0, padx=25, pady=5, sticky='w')

            reg_label = customtkinter.CTkLabel(self.labor_frame, text=f"Reg", font=self.bold_font)
            reg_label.grid(row=0, column=1, padx=(0,48), pady=5, sticky='e')
            
            ot_label = customtkinter.CTkLabel(self.labor_frame, text=f"OT", font=self.bold_font)
            ot_label.grid(row=0, column=2, padx=(0,50), pady=5, sticky='e')

            add_space_dt = max(cost_text_length - self.regular_font.measure('$0'), 0)
            dt_paddingx = 10
            dt_label = customtkinter.CTkLabel(self.labor_frame, text=f"DT", font=self.bold_font)
            dt_label.grid(row=0, column=3, padx=(0,dt_paddingx+add_space_dt), pady=5, sticky='e')

            add_space_cost = max(billed_text_length - self.regular_font.measure('$0'), 0)

            cost_paddingx = 5
            cost_label = customtkinter.CTkLabel(self.labor_frame, text=f"Labor Cost", font=self.bold_font)
            cost_label.grid(row=0, column=4, padx=(0,cost_paddingx+add_space_cost), pady=5, sticky='e')

            add_space_billed = max(profit_text_length - self.regular_font.measure('$0'), 0)

            billed_cost_label = customtkinter.CTkLabel(self.labor_frame, text=f"Billed Cost", font=self.bold_font)
            billed_cost_label.grid(row=0, column=5, padx=(0,15+add_space_billed), pady=5, sticky='e')
            
            add_space_profit = max(profit_percent_text_length - self.regular_font.measure('0%'), 0)
           
            profit_label = customtkinter.CTkLabel(self.labor_frame, text=f"Profit", font=self.bold_font)
            profit_label.grid(row=0, column=6, padx=(0,18+add_space_profit), pady=5, sticky='e')

            profit_percent_label = customtkinter.CTkLabel(self.labor_frame, text=f"Profit %", font=self.bold_font)
            profit_percent_label.grid(row=0, column=7, padx=(0,28), pady=5, sticky='e')


        elif maintenance_or_construction =='Construction' or project_cumulative == 'Cumulative Basis':

            self.scrollable_labor = ScrollablePanelFrame(self.labor_frame, corner_radius=0, border_width=1, border_color='#777')
            self.scrollable_labor.grid(row=1, column=0, columnspan=5, padx=(0,5), pady=(0,0), sticky="nsew")

            if project_cumulative == 'Cumulative Basis':
                if maintenance_or_construction =='Construction':
                    visits = self.work_order_data['Visits']['Construction']
                else:
                    if self.data_seg_button_1.get() == 'Maintenance':
                        visits = self.work_order_data['Visits']['Maintenance']
                    else:
                        visits = self.work_order_data['Visits']['Emergency Service']
            else:
                visits = self.work_order_data['Visits']

            visits = sorted(visits, key=lambda x: x['Date'], reverse=True)

            dates_list = {}

            for visit in visits:
                year = visit['Year']
                month = visit['Month']
                date = f'{month}_{year}'

                if (self.selected_timeframe == 'All Time' or (self.selected_timeframe == 'Yearly Basis' and self.selected_time == year)):

                    mechanics = visit['Mechanics']
                    reg = 0
                    ot = 0
                    dt = 0
                    cost = 0
                    for mechanic, hours in mechanics.items():
                        try:
                            hourly_rate = mechanic_settings[mechanic]
                        except:
                            hourly_rate = self.regular_hourly
                        reg += hours['Reg']
                        ot += hours['OT']
                        dt += hours['DT']
                        cost += ((hours['Reg'] * hourly_rate) + (hours['OT'] * hourly_rate * 1.5) + (hours['DT'] * hourly_rate * 2))
                    cost = round(cost)
                    reg_hours_list.append(reg); ot_hours_list.append(ot); dt_hours_list.append(dt); cost_list.append(cost)

                    if date not in [*dates_list.keys()]:
                        dates_list.update({date:{'Reg':reg,'OT':ot,'DT':dt,'Cost':cost}})
                    else:
                        dates_list[date]['Reg'] += reg
                        dates_list[date]['OT'] += ot
                        dates_list[date]['DT'] += dt
                        dates_list[date]['Cost'] += cost

            i = 0
            for date, hours in dates_list.items():
                month = date.split('_')[0]
                year = date.split('_')[1]
                reg = hours['Reg']
                ot = hours['OT']
                dt = hours['DT']
                cost = hours['Cost']

                if self.selected_timeframe == 'All Time':
                    year_text = f" '{year[-2:]}"
                else: 
                    year_text = ''               

                label = customtkinter.CTkLabel(self.scrollable_labor, text=f"{months_dictionary[month]}{year_text}", font=self.regular_font)
                label.grid(row=i, column=0, padx=25, pady=3, sticky='w')

                reg_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{reg}", font=self.regular_font)
                reg_hours.grid(row=i, column=1, padx=25, pady=3, sticky='e')

                ot_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{ot}", font=self.regular_font)
                ot_hours.grid(row=i, column=2, padx=25, pady=3, sticky='e')

                dt_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"{dt}", font=self.regular_font)
                dt_hours.grid(row=i, column=3, padx=25, pady=3, sticky='e')

                cost_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(cost))))]))
                cost_comma_str = f'-{cost_comma_str}' if cost < 0 else cost_comma_str
                cost_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"${cost_comma_str}", font=self.regular_font)
                cost_hours.grid(row=i, column=4, padx=(25,30), pady=3, sticky='e')      

                cost_text_length = max(len(f'${cost_comma_str}'), cost_text_length)            

                i += 1

            if i == 0:
                label = customtkinter.CTkLabel(self.scrollable_labor, text='EMPTY', font=self.regular_font)
                label.grid(row=i, column=0, padx=25, pady=4, sticky='w')

                reg_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0.0", font=self.regular_font)
                reg_hours.grid(row=i, column=1, padx=25, pady=3, sticky='e')

                ot_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0.0", font=self.regular_font)
                ot_hours.grid(row=i, column=2, padx=25, pady=3, sticky='e')

                dt_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"0.0", font=self.regular_font)
                dt_hours.grid(row=i, column=3, padx=(25,45), pady=3, sticky='e')

                cost_hours = customtkinter.CTkLabel(self.scrollable_labor, text=f"$0", font=self.regular_font)
                cost_hours.grid(row=i, column=4, padx=(25,30), pady=3, sticky='e')  

            month_label = customtkinter.CTkLabel(self.labor_frame, text=f"Month", font=self.bold_font)
            month_label.grid(row=0, column=0, padx=25, pady=5, sticky='w')

            reg_label = customtkinter.CTkLabel(self.labor_frame, text=f"Reg", font=self.bold_font)
            reg_label.grid(row=0, column=1, padx=(0,50), pady=5, sticky='e')

            ot_label = customtkinter.CTkLabel(self.labor_frame, text=f"OT", font=self.bold_font)
            ot_label.grid(row=0, column=2, padx=(0,50), pady=5, sticky='e')

            add_space_dt = int(cost_text_length / 2)
            dt_paddingx = 45 if project_cumulative == 'Cumulative Basis' else 35
            dt_label = customtkinter.CTkLabel(self.labor_frame, text=f"DT", font=self.bold_font)
            dt_label.grid(row=0, column=3, padx=(0,dt_paddingx+add_space_dt), pady=5, sticky='e')

            add_space_cost = 0

            cost_paddingx = 32
            cost_label = customtkinter.CTkLabel(self.labor_frame, text=f"Labor Cost", font=self.bold_font)
            cost_label.grid(row=0, column=4, padx=(0,35), pady=5, sticky='e')

        # ---------------------------- Bottom Summary ----------------------------

        total_label = customtkinter.CTkLabel(self.labor_frame, text=f"Total", font=self.bold_font)
        total_label.grid(row=2, column=0, padx=25, pady=5, sticky='w')

        total_reg_label = customtkinter.CTkLabel(self.labor_frame, text=f"{sum(reg_hours_list)}", font=self.bold_font)
        total_reg_label.grid(row=2, column=1, padx=(0,48), pady=5, sticky='e')

        total_ot_label = customtkinter.CTkLabel(self.labor_frame, text=f"{sum(ot_hours_list)}", font=self.bold_font)
        total_ot_label.grid(row=2, column=2, padx=(0,50), pady=5, sticky='e')

        total_dt_label = customtkinter.CTkLabel(self.labor_frame, text=f"{sum(dt_hours_list)}", font=self.bold_font)
        total_dt_label.grid(row=2, column=3, padx=(0,dt_paddingx+add_space_dt), pady=5, sticky='e')

        total_cost_sum = sum(cost_list)
        cost_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(total_cost_sum))))]))
        cost_comma_str = f'-{cost_comma_str}' if total_cost_sum < 0 else cost_comma_str
        total_cost_label = customtkinter.CTkLabel(self.labor_frame, text=f"${cost_comma_str}", font=self.bold_font)
        total_cost_label.grid(row=2, column=4, padx=(0,cost_paddingx+add_space_cost), pady=5, sticky='')

        if maintenance_or_construction == 'Maintenance' and project_cumulative == 'Project Basis':

            total_billed_sum = sum(billed_cost_list)
            billed_cost_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(total_billed_sum))))]))
            billed_cost_comma_str = f'-{billed_cost_comma_str}' if total_billed_sum < 0 else billed_cost_comma_str
            total_billed_cost_label = customtkinter.CTkLabel(self.labor_frame, text=f"${billed_cost_comma_str}", font=self.bold_font)
            total_billed_cost_label.grid(row=2, column=5, padx=(0,15+add_space_billed), pady=5, sticky='')

            total_profit_sum = sum(profit_list)
            profit_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(total_profit_sum))))]))
            profit_comma_str = f'-{profit_comma_str}' if total_profit_sum < 0 else profit_comma_str
            profit_label = customtkinter.CTkLabel(self.labor_frame, text=f"${profit_comma_str}", font=self.bold_font)
            profit_label.grid(row=2, column=6, padx=(0,18+add_space_profit), pady=5, sticky='')

            profit_percent = round(sum(profit_list) * 100 / sum(billed_cost_list),1) if sum(billed_cost_list) != 0 else 0
            profit_percent_label = customtkinter.CTkLabel(self.labor_frame, text=f"{profit_percent}%", font=self.bold_font)
            profit_percent_label.grid(row=2, column=7, padx=(0,28), pady=5, sticky='')

        self.labor_cost = sum(cost_list)
        self.labor_billed = sum(billed_cost_list)
        self.labor_reg = sum(reg_hours_list)
        self.labor_ot = sum(ot_hours_list)
        self.labor_dt = sum(dt_hours_list)
        self.total_visits = i

    def filter_panel_builder(self):
            
        try:
            project_settings = self.settings[self.update_jobs[self.previous_project_index]]
        except:
            project_settings = {}

        month_label = customtkinter.CTkLabel(self.filter_frame, text=f"Month", font=self.bold_font)
        month_label.grid(row=0, column=0, padx=25, pady=5, sticky='w')

        filter_label = customtkinter.CTkLabel(self.filter_frame, text=f"Filter Cost", font=self.bold_font)
        filter_label.grid(row=0, column=1, padx=(25, 40), pady=5, sticky='')

        cost_label = customtkinter.CTkLabel(self.filter_frame, text=f"Billed Cost", font=self.bold_font)
        cost_label.grid(row=0, column=2, padx=25, pady=5, sticky='')

        profit_label = customtkinter.CTkLabel(self.filter_frame, text=f"Profit", font=self.bold_font)
        profit_label.grid(row=0, column=3, padx=(25, 35), pady=5, sticky='')

        self.scrollable_filter = ScrollablePanelFrame(self.filter_frame, corner_radius=0, border_width=1, border_color='#777')
        self.scrollable_filter.grid(row=1, column=0, columnspan=4, padx=(0,5), pady=(0,0), sticky="nsew")

        visits = sorted(self.work_order_data['Visits'], key=lambda x: x['Date'], reverse=True)

        maintenance_text = ''
        cost_list = []
        billed_cost_list = []

        self.filter_entry_label_list = []
        self.filter_billed_label_list = []
        self.filter_date_list = []

        i = 0
        for visit in visits:
            year = visit['Year']
            month = visit['Month']
            date = visit['Date']

            if (self.selected_timeframe == 'All Time' or (self.selected_timeframe == 'Yearly Basis' and self.selected_time == year)) and (visit['Job_Type'] == self.selected_job_type):

                if self.selected_job_type == 'Maintenance':
                    letter = ''; interval = ''
                    if self.maintenance_frequency == 'quarterly':
                        letter = 'Q'
                        interval = ceil(round(int(month) / 3, 1))
                    elif self.maintenance_frequency == 'bi-monthly':
                        letter = 'B'
                        interval = ceil(round(int(month) / 2, 1))
                    elif self.maintenance_frequency == 'monthly':
                        letter = 'M'
                        interval = int(month)
                    elif self.maintenance_frequency == '2x/year':
                        letter = 'T'
                        interval = ceil(round(int(month) / 6, 1))
                    elif self.maintenance_frequency == '3x/year':
                        letter = 'T'
                        interval = ceil(round(int(month) / 4, 1))
                    if letter:
                        maintenance_text = f'{letter}{interval} - '
                    else:
                        maintenance_text = ''

                if self.selected_timeframe == 'All Time':
                    year_text = f" '{year[-2:]}"
                else: 
                    year_text = ''

                label = customtkinter.CTkLabel(self.scrollable_filter, text=f"{maintenance_text}{months_dictionary[month]}{year_text}", font=self.regular_font)
                label.grid(row=i, column=0, padx=25, pady=3, sticky='w')

                filter_price_entry = customtkinter.CTkEntry(self.scrollable_filter, width=90, font=self.regular_font)
                filter_price_entry.grid(row=i, column=1, padx=20, pady=3, sticky='')
                filter_price_entry.bind("<Return>", self.update_filter_prices)
                try:
                    filter_entry_value = project_settings['Filter_Entry'][date]
                    filter_entry_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(filter_entry_value))))]))
                    filter_entry_text = f'-{filter_entry_text}' if filter_entry_value < 0 else filter_entry_text
                    filter_price_entry.insert(0, f"{filter_entry_text}")
                except:
                    filter_price_entry.insert(0, "0")
                self.filter_entry_label_list.append(filter_price_entry)
                cost_list.append(round(float(filter_price_entry.get().replace(',','') or 0)))

                filter_billed_entry = customtkinter.CTkEntry(self.scrollable_filter, width=90, font=self.regular_font)
                filter_billed_entry.grid(row=i, column=2, padx=20, pady=3, sticky='')
                filter_billed_entry.bind("<Return>", self.update_filter_prices)
                try:
                    filter_billed_value = project_settings['Filter_Billed'][date]
                    filter_billed_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(filter_billed_value))))]))
                    filter_billed_text = f'-{filter_billed_text}' if filter_billed_value < 0 else filter_billed_text
                    filter_billed_entry.insert(0, f"{filter_billed_text}")
                except:
                    filter_billed_entry.insert(0, "0")
                self.filter_billed_label_list.append(filter_billed_entry)
                billed_cost_list.append(round(float(filter_billed_entry.get().replace(',','') or 0)))

                profit_i = round(billed_cost_list[-1] - cost_list[-1])

                profit_i_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(profit_i))))]))
                profit_i_text = f'-{profit_i_text}' if profit_i < 0 else profit_i_text
                profit_hours = customtkinter.CTkLabel(self.scrollable_filter, text=f"${profit_i_text}", font=self.regular_font)
                profit_hours.grid(row=i, column=3, padx=25, pady=3, sticky='')
                self.filter_date_list.append(date)

                i += 1

        if i == 0:
            label = customtkinter.CTkLabel(self.scrollable_filter, text='EMPTY', font=self.regular_font)
            label.grid(row=i, column=0, padx=25, pady=4, sticky='w')


        total_label = customtkinter.CTkLabel(self.filter_frame, text=f"Total", font=self.bold_font)
        total_label.grid(row=2, column=0, padx=25, pady=5, sticky='w')

        cost = sum(cost_list)
        cost_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(cost))))]))
        cost_comma_str = f'-{cost_comma_str}' if cost < 0 else cost_comma_str
        self.total_filter_price_label = customtkinter.CTkLabel(self.filter_frame, text=f"${cost_comma_str}", font=self.bold_font)
        self.total_filter_price_label.grid(row=2, column=1, padx=25, pady=5, sticky='')

        billed_cost = sum(billed_cost_list)
        billed_comma_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(billed_cost))))]))
        billed_comma_str = f'-{billed_comma_str}' if billed_cost < 0 else billed_comma_str
        self.total_filter_billed_label = customtkinter.CTkLabel(self.filter_frame, text=f"${billed_comma_str}", font=self.bold_font)
        self.total_filter_billed_label.grid(row=2, column=2, padx=25, pady=5, sticky='')

        profit = round(sum(billed_cost_list) - sum(cost_list))
        profit_str = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(profit))))]))
        profit_str = f'-{profit_str}' if profit < 0 else profit_str
        total_filter_profit_label = customtkinter.CTkLabel(self.filter_frame, text=f"${profit_str}", font=self.bold_font)
        total_filter_profit_label.grid(row=2, column=3, padx=(25, 50), pady=5, sticky='')

        self.filter_cost = sum(cost_list)
        self.filter_billed = sum(billed_cost_list)

    def profit_panel_builder(self):
        # ROI 3
        maintenance_or_construction = self.seg_button_1.get()
        project_cumulative = self.seg_button_2.get()
        self.labor_panel_builder()
        if self.scrollable_labor:
            for widget in self.scrollable_labor.winfo_children(): 
                widget.destroy()
        for widget in self.labor_frame.winfo_children(): widget.destroy()
        self.scrollable_labor = None

        if maintenance_or_construction == 'Maintenance' and project_cumulative == 'Project Basis':

            if self.selected_job_type == 'Maintenance':
                self.filter_panel_builder()
                if self.scrollable_filter:
                    for widget in self.scrollable_filter.winfo_children(): 
                        widget.destroy()
                for widget in self.filter_frame.winfo_children(): widget.destroy()
                self.scrollable_filter = None
                extra_text = ' + Filters)'
                padx_int = 20
            else:
                self.filter_cost = 0
                self.filter_billed = 0
                extra_text = ')'
                padx_int = 71

            total_cost_label = customtkinter.CTkLabel(self.profit_frame, text=f"PROFITABILITY", font=self.bold_font)
            total_cost_label.grid(row=0, column=0, padx=(90,20), pady=5, sticky='w')

            self.profit_frame1 = customtkinter.CTkFrame(self.profit_frame, corner_radius=0, border_width=1, border_color='#777')
            self.profit_frame1.grid(row=1, column=0, padx=0, pady=(0, 0), sticky="nsew")
            self.profit_frame1.grid_rowconfigure(0, weight=1)
            self.profit_frame1.grid_columnconfigure(1, weight=1)

            total_cost_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Cost (Labor{extra_text}:", font=self.regular_font)
            total_cost_label.grid(row=0, column=0, padx=20, pady=5, sticky='e')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Billed (Labor{extra_text}:", font=self.regular_font)
            reg_label.grid(row=3, column=0, padx=(padx_int,20), pady=5, sticky='e')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Profit:", font=self.regular_font)
            reg_label.grid(row=4, column=0, padx=20, pady=5, sticky='e')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Profit Percentage:", font=self.regular_font)
            reg_label.grid(row=5, column=0, padx=20, pady=5, sticky='e')

            cost = round(self.labor_cost + self.filter_cost)
            cost_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(cost))))]))
            cost_text = f'-{cost_text}' if cost < 0 else cost_text
            total_cost_label = customtkinter.CTkLabel(self.profit_frame1, text=f"${cost_text}", font=self.regular_font)
            total_cost_label.grid(row=0, column=1, padx=20, pady=5, sticky='w')

            billed = round(self.labor_billed + self.filter_billed)
            billed_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(billed))))]))
            billed_text = f'-{billed_text}' if billed < 0 else billed_text
            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"${billed_text}", font=self.regular_font)
            reg_label.grid(row=3, column=1, padx=20, pady=5, sticky='w')

            profit = round(billed-cost)
            profit_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(profit))))]))
            profit_text = f'-{profit_text}' if profit < 0 else profit_text
            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"${profit_text}", font=self.regular_font)
            reg_label.grid(row=4, column=1, padx=20, pady=5, sticky='w')

            profit_percent = round(profit * 100 / billed, 1) if billed != 0 else 0
            text_var = '#137f34' if profit_percent >= 0 else '#ca0b0b'
            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"{profit_percent}%", text_color=text_var, font=self.regular_font)
            reg_label.grid(row=5, column=1, padx=20, pady=5, sticky='w')

        else:

            total_cost_label = customtkinter.CTkLabel(self.profit_frame, text=f"COST METRICS", font=self.bold_font)
            total_cost_label.grid(row=0, column=0, padx=(90,20), pady=5, sticky='w')

            self.profit_frame1 = customtkinter.CTkFrame(self.profit_frame, corner_radius=0, border_width=1, border_color='#777')
            self.profit_frame1.grid(row=1, column=0, padx=0, pady=(0, 0), sticky="nsew")
            self.profit_frame1.grid_rowconfigure(0, weight=1)
            self.profit_frame1.grid_columnconfigure(1, weight=1)

            total_cost_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Cost (Labor):", font=self.regular_font)
            total_cost_label.grid(row=0, column=0, padx=(75,20), pady=5, sticky='e')

            cost_text = "".join(reversed([digit + ',' if index % 3 == 0 and index != 0 else digit for index, digit in enumerate(reversed(str(abs(self.labor_cost))))]))
            cost_text = f'-{cost_text}' if self.labor_cost < 0 else cost_text
            total_cost_label = customtkinter.CTkLabel(self.profit_frame1, text=f"${cost_text}", font=self.regular_font)
            total_cost_label.grid(row=0, column=1, padx=20, pady=5, sticky='w')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Hours (Reg):", font=self.regular_font)
            reg_label.grid(row=2, column=0, padx=20, pady=5, sticky='e')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f" {self.labor_reg}", font=self.regular_font)
            reg_label.grid(row=2, column=1, padx=20, pady=5, sticky='w')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Hours (OT):", font=self.regular_font)
            reg_label.grid(row=3, column=0, padx=20, pady=5, sticky='e')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f" {self.labor_ot}", font=self.regular_font)
            reg_label.grid(row=3, column=1, padx=20, pady=5, sticky='w')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f"Total Hours (DT):", font=self.regular_font)
            reg_label.grid(row=4, column=0, padx=20, pady=5, sticky='e')

            reg_label = customtkinter.CTkLabel(self.profit_frame1, text=f" {self.labor_dt}", font=self.regular_font)
            reg_label.grid(row=4, column=1, padx=20, pady=5, sticky='w')


    def personel_panel_builder(self):
        name_label = customtkinter.CTkLabel(self.personel_frame, text=f"Name", font=self.bold_font)
        name_label.grid(row=0, column=0, padx=25, pady=5, sticky='w')

        months_label = customtkinter.CTkLabel(self.personel_frame, text=f"Month", font=self.bold_font)
        months_label.grid(row=0, column=1, padx=20, pady=5, sticky='')

        reg_label = customtkinter.CTkLabel(self.personel_frame, text=f"Reg", font=self.bold_font)
        reg_label.grid(row=0, column=2, padx=(27,24), pady=5, sticky='')

        ot_label = customtkinter.CTkLabel(self.personel_frame, text=f"OT", font=self.bold_font)
        ot_label.grid(row=0, column=3, padx=25, pady=5, sticky='')

        dt_label = customtkinter.CTkLabel(self.personel_frame, text=f"DT", font=self.bold_font)
        dt_label.grid(row=0, column=4, padx=(26, 47), pady=5, sticky='')

        self.scrollable_personel = ScrollablePanelFrame(self.personel_frame, corner_radius=0, border_width=1, border_color='#777')
        self.scrollable_personel.grid(row=1, column=0, columnspan=5, padx=(0,5), pady=(0,0), sticky="nsew")

        visits = sorted(self.work_order_data['Visits'], key=lambda x: x['Date'], reverse=True)

        maintenance_or_construction = self.seg_button_1.get()

        i = 0
        for visit in visits:
            year = visit['Year']
            month = visit['Month']

            if (self.selected_timeframe == 'All Time' or (self.selected_timeframe == 'Yearly Basis' and self.selected_time == year)) and (visit['Job_Type'] == self.selected_job_type or maintenance_or_construction == 'Construction'):
               
                if self.selected_timeframe == 'All Time':
                    year_text = f" '{year[-2:]}"
                else: 
                    year_text = '   '
            
                mechanics = visit['Mechanics']
                for mechanic, hours in mechanics.items():

                    label = customtkinter.CTkLabel(self.scrollable_personel, text=mechanic, font=self.regular_font)
                    label.grid(row=i, column=0, padx=25, pady=4, sticky='w')

                    months_name = customtkinter.CTkLabel(self.scrollable_personel, text=f"{months_abv_dictionary[month]}{year_text}", font=self.regular_font)
                    months_name.grid(row=i, column=1, padx=25, pady=4, sticky='')

                    reg_hours = customtkinter.CTkLabel(self.scrollable_personel, text=hours['Reg'], font=self.regular_font)
                    reg_hours.grid(row=i, column=2, padx=25, pady=4, sticky='')

                    ot_hours = customtkinter.CTkLabel(self.scrollable_personel, text=hours['OT'], font=self.regular_font)
                    ot_hours.grid(row=i, column=3, padx=25, pady=4, sticky='')

                    dt_hours = customtkinter.CTkLabel(self.scrollable_personel, text=hours['DT'], font=self.regular_font)
                    dt_hours.grid(row=i, column=4, padx=25, pady=4, sticky='')

                    i += 1

        if i == 0:
            label = customtkinter.CTkLabel(self.scrollable_personel, text='EMPTY', font=self.regular_font)
            label.grid(row=i, column=0, padx=25, pady=4, sticky='w')

    # ---------- Buttons ------------

    def select_timeframe_event(self, item, index):
        self.selected_time = item
        self.time_label_list[self.previous_time_index].configure(fg_color='transparent', text_color='gray10')

        self.time_label_list[index].configure(fg_color='#333', text_color='#f5f5f5')
 
        self.previous_time_index = index
        self.panel_builder(param=self.data_seg_button_2.get())

    def timeframe_update(self):
        self.years_list = sorted([*self.work_order_data['Years'].keys()], reverse=True)

        self.previous_time_index = 0
     
        for i in range(len(self.years_list)):
            self.time_frame_reconfigure(item=self.years_list[i], index=i)
        length = int(len(self.time_label_list))
        for i in range(len(self.years_list),length):
            self.time_frame_destroy_element()

        self.time_seg_button(param=self.data_seg_button_3.get())

    def select_project_event(self, item, index):
        error_tuple = ()
        self.scrollable_label_button_frame.label_list[self.previous_project_index].configure(fg_color='transparent', text_color='gray10')
      
        self.project_label.configure(text=f"PROJECT: {item.upper()}")

        self.scrollable_label_button_frame.label_list[index].configure(fg_color='#333', text_color='#f5f5f5')
        self.previous_project_index = index

        self.page_project = (self.project_page_index, index)

        if self.seg_button_1.get() == 'Maintenance':
            job_number = item[:6]
            self.job_number_label.configure(text=f"JOB #: {job_number}")
            try:
                job_data = self.DataAnalysis.maintenance_database_clients[job_number]
                name = job_data['Name'] if job_data['Name'] is not None else 'EXCEL BLANK'
                address = job_data['Address'] if job_data['Address'] is not None else 'EXCEL BLANK'
                frequency = job_data['Frequency'] if job_data['Frequency'] is not None else 'EXCEL BLANK'
                service_rate = job_data['Service_Rate'] if job_data['Service_Rate'] is not None else 'EXCEL BLANK'
                self.maintenance_frequency = frequency
                self.cost_per_visit = int(job_data['Cost_Per_Visit'] or 0)
                self.service_hourly = int(job_data['Service_Rate'] or 0)
                # Update Labels
                self.name_label.configure(text=f"NAME: {name.upper()}")
                self.address_label.configure(text=f"ADDRESS: {address.upper()}")
                self.pm_frequency_label.configure(text=f"PM FREQUENCY: {frequency.upper()}")
                self.service_cost_label.configure(text=f"EMERGENCY BILLABLE ($ / Hr): {service_rate}")
                if 'EXCEL BLANK' in [name, address, frequency, service_rate] or self.cost_per_visit == 0:
                    error_tuple = ('Excel File Error', 'Project Data Not in Maintenance Excel File')
            except:
                self.name_label.configure(text=f"NAME: EXCEL BLANK")
                self.address_label.configure(text=f"ADDRESS: EXCEL BLANK")
                self.pm_frequency_label.configure(text=f"PM FREQUENCY: EXCEL BLANK")
                self.service_cost_label.configure(text=f"EMERGENCY BILLABLE ($ / Hr): EXCEL BLANK")
                self.maintenance_frequency = 'None'
                self.cost_per_visit = 0
                self.service_hourly = 0
                error_tuple = ('Excel File Error', 'Project Data Not in Maintenance Excel File')
                
            path = f'S:/PIERPONT MAINTENANCE & SERVICE/Data_Storage/File_System_Scans/Maintenance/{self.scanned_jobs[self.update_jobs[index]]}.json'
            try:
                with open(path, 'r') as j:
                   self.work_order_data = json.load(j)
            except:
                 self.work_order_data = {'Years': {},'Visits': []}
                 #error_tuple = ('Data Error', 'JSON File Location Missing. Work Orders may not be present.')
        else:
            path = f'S:/PIERPONT MAINTENANCE & SERVICE/Data_Storage/File_System_Scans/Construction/{self.scanned_jobs[self.update_jobs[index]]}.json'
            try:
                with open(path, 'r') as j:
                    work_order_data = json.load(j)
                if 'Name' in [*work_order_data.keys()]:
                    self.name_label.configure(text=f"NAME: {work_order_data['Name'].upper()}")
                    self.address_label.configure(text=f"ADDRESS: {work_order_data['Address'].upper()}")
                    if not work_order_data['Job#']:
                        job_number = item.split('_')[0]
                        job_number = f'{job_number[3:]}-{job_number[:2]}'
                    else:
                        job_number = work_order_data['Job#']
                    self.job_number_label.configure(text=f"JOB #: {job_number}")
                else:
                    self.name_label.configure(text=f"NAME: EMPTY FILE")
                    self.address_label.configure(text=f"ADDRESS: EMPTY FILE")
                    job_number = item.split('_')[0]
                    job_number = f'{job_number[3:]}-{job_number[:2]}'
                    self.job_number_label.configure(text=f"JOB #: {job_number}")
                self.work_order_data = work_order_data
            except:
                self.name_label.configure(text=f"NAME: FAILED LOADING")
                self.address_label.configure(text=f"ADDRESS: FAILED LOADING")
                self.job_number_label.configure(text=f"JOB #: 000-00")
                self.work_order_data = {'Years': {},'Visits': []}
                #error_tuple = ('Data Error', 'JSON File Location Missing. Work Orders may not be present.')
        self.timeframe_update()
        if error_tuple:
            messagebox.showerror(error_tuple[0], error_tuple[1])


    def select_cumulative_event(self):
        error_tuple = ()
        if self.seg_button_1.get() == 'Maintenance':
            self.project_label.configure(text=f"PROJECT: FULL MAINTENANCE ANALYSIS")
            path = f'S:/PIERPONT MAINTENANCE & SERVICE/Data_Storage/File_System_Scans/Maintenance/full_maintenance.json'
        else:
            self.project_label.configure(text=f"PROJECT: FULL CONSTRUCTION ANALYSIS")
            path = f'S:/PIERPONT MAINTENANCE & SERVICE/Data_Storage/File_System_Scans/Construction/full_construction.json'
        try:
            with open(path, 'r') as j:
                work_order_data = json.load(j)
            self.name_label.configure(text=f"NAME: {work_order_data['Name'].upper()}")
            self.address_label.configure(text=f"ADDRESS: {work_order_data['Address'].upper()}")    
            self.job_number_label.configure(text=f"JOB #: NONE")
            self.work_order_data = work_order_data
        except:
            self.name_label.configure(text=f"NAME: FAILED LOADING")
            self.address_label.configure(text=f"ADDRESS: FAILED LOADING")
            self.job_number_label.configure(text=f"JOB #: 000-00")
            self.work_order_data = {'Years': {},'Visits': []}
            error_tuple = ('Data Error', 'JSON File Location Missing.')
        self.timeframe_update()
        if error_tuple:
            messagebox.showerror(error_tuple[0], error_tuple[1])

    def project_next_button(self):
        page_ratio = (len(self.scanned_jobs) / 20)
        if self.project_page_index < page_ratio:
            self.project_page_index += 1
            self.project_navigation_update()

    def project_previous_button(self):
        if self.project_page_index > 1:
            self.project_page_index -= 1
            self.project_navigation_update()

    def project_navigation_update(self):
        self.page_label.configure(text=f'Page {self.project_page_index}')

        self.previous_project_index = 0

        if self.seg_button_2.get() == 'Project Basis':
            self.scrollable_label_button_frame.grid(row=4, column=5, rowspan=3, padx=(10,20), pady=(0, 20), sticky="nsew")
        else:
            self.scrollable_label_button_frame.grid_forget()

        if self.seg_button_1.get() == 'Maintenance':
            job_names = sorted([*self.scanned_jobs.keys()])
        else:
            job_names = sorted([*self.scanned_jobs.keys()], reverse=True)
        update_jobs = job_names[((self.project_page_index-1)*20):(self.project_page_index*20)]
        self.update_jobs = update_jobs
   
        for i in range(len(update_jobs)): 
            self.scrollable_label_button_frame.reconfigure(update_jobs[i].upper(), i)
        for i in range(len(update_jobs),20):
            self.scrollable_label_button_frame.destroy_element()

        if self.project_page_index == self.page_project[0]:
            self.scrollable_label_button_frame.label_list[self.page_project[1]].configure(fg_color='#333', text_color='#f5f5f5')
            self.previous_project_index = self.page_project[1]
           

    # ---------- Segmented Buttons ------------

    def maintenance_construction_navigation(self, param: str):
        maintenace_construction = self.seg_button_1.get()
        project_cumulative = self.seg_button_2.get()
        if maintenace_construction == "Maintenance" and project_cumulative == 'Project Basis':
            self.data_seg_button_1.grid(row=3, column=0, columnspan=2, ipadx=15, padx=20, pady=(0, 10), sticky="w")
            self.data_seg_button_2.grid(row=4, column=0, columnspan=2, ipadx=20, padx=20, pady=(0,10), sticky="w")
            if self.data_seg_button_1.get() == 'Maintenance':
                self.pm_frequency_label.grid(row=2, column=1, padx=15, pady=(0,5), sticky='w')
            else:
                self.service_cost_label.grid(row=2, column=1, padx=15, pady=(0,5), sticky='w')
            self.panel_row = 5
            self.panel_rowspan = 1               
            self.data_seg_button_2.configure(values=["Mechanic Name","Filter Cost","Labor Cost","Project Summary"])
            self.scanned_jobs = self.DataAnalysis.scanned_maintenance_jobs
            self.project_page_index = 1
            self.scrollable_label_button_frame.grid(row=4, column=5, columnspan=2, padx=(10,20), pady=(0, 10), sticky="nsew")
            self.page_frame.grid(row=3, column=5, columnspan=2, padx=(10,20), pady=(10, 0), sticky="nsew")
        elif maintenace_construction == 'Maintenance' and project_cumulative == 'Cumulative Basis':
            self.data_seg_button_1.grid(row=3, column=0, columnspan=2, ipadx=15, padx=20, pady=(0, 10), sticky="w")
            self.data_seg_button_2.grid(row=4, column=0, columnspan=2, ipadx=20, padx=20, pady=(0,10), sticky="w")
            self.pm_frequency_label.grid_forget()
            self.service_cost_label.grid_forget()
            self.panel_row = 5
            self.panel_rowspan = 1    

            if self.data_seg_button_2.get() in ['Filter Cost', 'Mechanic Name']:
                self.data_seg_button_2.set("Project Summary") 
            self.data_seg_button_2.configure(values=["Labor Cost","Project Summary"])
            self.page_frame.grid_forget()
            self.scrollable_label_button_frame.grid_forget()
        elif maintenace_construction == 'Construction' and project_cumulative == 'Project Basis':
            self.data_seg_button_1.grid_forget()
            self.pm_frequency_label.grid_forget()
            self.service_cost_label.grid_forget()
            self.data_seg_button_2.grid(row=3, column=0, columnspan=2, ipadx=20, padx=20, pady=(0,10), sticky="w")
            self.panel_row = 4
            self.panel_rowspan = 2
            self.data_seg_button_2.configure(values=["Mechanic Name","Labor Cost","Project Summary"])
     
            if self.data_seg_button_2.get() == 'Filter Cost':
                self.data_seg_button_2.set("Project Summary")
            self.scanned_jobs = self.DataAnalysis.scanned_construction_jobs
            self.project_page_index = 1
            self.scrollable_label_button_frame.grid(row=4, column=5, columnspan=2, padx=(10,20), pady=(0, 10), sticky="nsew")
            self.page_frame.grid(row=3, column=5, columnspan=2, padx=(10,20), pady=(10, 0), sticky="nsew")
        else:
            self.data_seg_button_1.grid_forget()
            self.pm_frequency_label.grid_forget()
            self.service_cost_label.grid_forget()
            self.data_seg_button_2.grid(row=3, column=0, columnspan=2, ipadx=20, padx=20, pady=(0,10), sticky="w")
            self.panel_row = 4
            self.panel_rowspan = 2
            self.data_seg_button_2.configure(values=["Labor Cost","Project Summary"])
       
            if self.data_seg_button_2.get() in ['Filter Cost', 'Mechanic Name']:
                self.data_seg_button_2.set("Project Summary")
            self.page_frame.grid_forget()
            self.scrollable_label_button_frame.grid_forget()
            
        if self.seg_button_2.get() == 'Project Basis':
            self.project_navigation_update()
            if self.update_jobs:
                self.select_project_event(index=0,item=self.update_jobs[0])
        else:
            self.select_cumulative_event()     

    def job_type_seg_button(self, param: str):
        self.selected_job_type = param
        project_cumulative = self.seg_button_2.get()
        if param == 'Maintenance' and project_cumulative == 'Project Basis':
            self.pm_frequency_label.grid(row=2, column=1, padx=15, pady=(0,5), sticky='w')
            self.service_cost_label.grid_forget()
            self.data_seg_button_2.configure(values=["Mechanic Name","Filter Cost","Labor Cost","Project Summary"])
        elif param == 'Emergency Service' and project_cumulative == 'Project Basis':
            self.data_seg_button_2.configure(values=["Mechanic Name","Labor Cost","Project Summary"])
            self.service_cost_label.grid(row=2, column=1, padx=15, pady=(0,5), sticky='w')
            self.pm_frequency_label.grid_forget()
            selected_seg_button_2 = self.data_seg_button_2.get()
            if selected_seg_button_2 == 'Filter Cost':
                self.data_seg_button_2.set("Project Summary")
        self.panel_builder(param=self.data_seg_button_2.get())

    # ---------- Time Segmented Buttons ------------

    def time_seg_button(self, param: str):
        self.selected_timeframe = param
        if param == 'All Time':
            self.time_frame.grid_forget()
            self.panel_builder(param=self.data_seg_button_2.get())
        else:
            self.time_frame.grid(row=4, column=2, rowspan=2, padx=(10,20), pady=(0, 20), sticky="nsew")
            if self.years_list:
                self.select_timeframe_event(item=self.years_list[0], index=0)
            else:
                self.panel_builder(param=self.data_seg_button_2.get())

    #  ------------ Time Frame ------------

    def time_frame_add_item(self, item: str, index: int):
        label = customtkinter.CTkLabel(self.time_frame, text=item, compound="left", padx=5, corner_radius=5, font=customtkinter.CTkFont(family='Roboto', size=13, weight="bold"))
        button = customtkinter.CTkButton(self.time_frame, text="Select", width=60, height=20, font=customtkinter.CTkFont(family='Roboto', size=12, weight="normal"))

        button.configure(command=lambda: self.select_timeframe_event(item, index))
        label.grid(row=index, column=0, ipady=0, ipadx=5, pady=(8,0), padx=(15, 0), sticky="w")
        button.grid(row=index, column=1, pady=(8,0), padx=(5,15), sticky='e')
        self.time_label_list.append(label)
        self.time_button_list.append(button)

    def time_frame_reconfigure(self, item: str, index: int):
        if index > len(self.time_label_list) - 1:
            self.time_frame_add_item(item, index)
        else:
            self.time_label_list[index].configure(text=item, fg_color='transparent', text_color='gray10')
            self.time_button_list[index].configure(command=lambda: self.select_timeframe_event(item, index))
       
    def time_frame_destroy_element(self):
        self.time_label_list[-1].destroy()
        self.time_button_list[-1].destroy()
        del self.time_label_list[-1], self.time_button_list[-1]
    

    # UTILITIES
    def activate_refresh(self):
        if self.toplevel_window is not None:
            mechanic_settings = {'Mechanic Hourly Rates': self.toplevel_window.settings['Mechanic Hourly Rates']}
            self.settings = self.settings | mechanic_settings
        self.recalculate_regular_price(focus=False)
        if self.data_seg_button_2.get() == 'Filter Cost':
            self.update_filter_prices()

    def update_filter_prices(self, entry_field = ''):
        self.main_frame.focus()
        error = 0
        for i in range(len(self.filter_entry_label_list)):
            date = self.filter_date_list[i]
            value_entry = self.filter_entry_label_list[i].get().replace(',','')
            value_billed = self.filter_billed_label_list[i].get().replace(',','')
            try:
                float_entry = round(float(value_entry or 0))
                try:
                    self.settings[self.update_jobs[self.previous_project_index]]
                except:
                    self.settings.update({self.update_jobs[self.previous_project_index]:{'Filter_Entry':{},'Filter_Billed':{}}})

                self.settings[self.update_jobs[self.previous_project_index]]['Filter_Entry'].update({date:float_entry})
            except:
                error += 1
                self.filter_entry_label_list[i].delete(0, len(value_entry))
            try:
                float_billed = round(float(value_billed or 0))
                try:
                    self.settings[self.update_jobs[self.previous_project_index]]
                except:
                    self.settings.update({self.update_jobs[self.previous_project_index]:{'Filter_Entry':{},'Filter_Billed':{}}})

                self.settings[self.update_jobs[self.previous_project_index]]['Filter_Billed'].update({date:float_billed})
            except:
                error += 1
                self.filter_billed_label_list[i].delete(0, len(value_billed))

        self.panel_builder(param=self.data_seg_button_2.get())
        if error:
            messagebox.showerror('Number Input Error', 'Please Enter Number Only.')

    def recalculate_regular_price(self, entry_field = '', update_panel = True, focus = True):
        if focus: self.main_frame.focus()
        try:
            self.regular_hourly = int(self.reg_price_entry.get() or 0)
        except:
            self.reg_price_entry.delete(0, len(self.reg_price_entry.get()))
            self.reg_price_entry.insert(0, f'{self.regular_hourly}')
            messagebox.showerror('Number Input Error', 'Please Enter Number Only.')
        if update_panel:
            self.panel_builder(param=self.data_seg_button_2.get())

    def open_toplevel(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ToplevelWindow(self)  # create window if its None or destroyed
            self.after(100, self.toplevel_window.focus)
            self.toplevel_window.apply_settings(settings=self.settings)
        else:
            self.toplevel_window.focus()  # if window exists focus it

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def save_settings(self):
        path = "S:/PIERPONT MAINTENANCE & SERVICE/Data_Storage/File_System_Scans/pierpont_analytics_settings.json" 
        try:
            with open(path, 'w') as j:
                json.dump(self.settings, j)
        except:
            pass

        self.destroy()

    def save_as(self):
        # Ask for Directory
        file_to_print = filedialog.askopenfilename(
        initialdir="/", title="Select file", 
        filetypes=(("Microsoft Excel Worksheet", "*.xlsx")))

    def print_file(self):
        x_pos, y_pos = self.winfo_x() + 8, self.winfo_y() + 2
        width, height = self.winfo_width(), self.winfo_height() + 29

        pyautogui.screenshot(imageFilename='screenshot.jpg', region=(x_pos, y_pos, width, height))
        # Print Hard Copy of File
        win32api.ShellExecute(0, "print", 'screenshot.jpg', None, ".", 0)
        
# Project Scrollable
class ScrollableLabelButtonFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)

        self.command = command
        self.label_list = [] 
        self.button_list = []

    def add_item(self, item, index: int):
        label = customtkinter.CTkLabel(self, text=item, compound="left", corner_radius=6, padx=5, anchor="w", font = customtkinter.CTkFont(family='Roboto', size=13, weight="bold"))
        button = customtkinter.CTkButton(self, text="Select", width=70, height=20, font=customtkinter.CTkFont(family='Roboto', size=12, weight="normal"))
        if self.command is not None:
            button.configure(command=lambda: self.command(item, index))
        label.grid(row=index, column=0, pady=6, ipady=1, padx=(10, 0), sticky="w")
        button.grid(row=index, column=1, pady=6, padx=(10,5))
        self.label_list.append(label)
        self.button_list.append(button)

    def reconfigure(self, item, index: int):
        if index > len(self.label_list) - 1:
            self.add_item(item, index)
        else:
            self.label_list[index].configure(text=item, fg_color='transparent', text_color='gray10')
            self.button_list[index].configure(command=lambda: self.command(item, index))

    def destroy_element(self):
        self.label_list[-1].destroy()
        self.button_list[-1].destroy()
        del self.label_list[-1], self.button_list[-1]
 
# Panel Scrollable
class ScrollablePanelFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)

# Setting Window
class ToplevelWindow(customtkinter.CTkToplevel):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(icon_index=1, *args, **kwargs)
        
        self.parent = parent
        # Window Title
        self.title("Pierpont Mechanical Business Analytics Settings")

        # Default Window Size
        self.geometry("750x750")

        # Configure App Grid Layout (4x4)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        title_font = customtkinter.CTkFont(family='Alegreya SC', size=20, weight="bold")
        self.regular_font = customtkinter.CTkFont(family='Roboto', size=14, weight="normal")
        self.large_bold = customtkinter.CTkFont(family='Roboto', size=16, weight="bold")
        self.button_font = customtkinter.CTkFont(family='Roboto', size=13, weight="normal")
        self.bold_font = customtkinter.CTkFont(family='Roboto', size=14, weight="bold")

        """ 
            Header Frame 
        """
        self.header_frame = customtkinter.CTkFrame(self, height=100, corner_radius=0, fg_color=('#700e10','#333333'))
        self.header_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self.header_frame.grid_rowconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=1)
        # App Title
        title_label = customtkinter.CTkLabel(self.header_frame, text="PIERPONT MECHANICAL SETTINGS",corner_radius=6, text_color='#f5f5f5', font=title_font)
        title_label.grid(row=0, column=0, ipadx=10, ipady=1, padx=25, pady=10, sticky='n')

        """ 
            Main Frame 
        """
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=6, border_width=1, border_color='#777')
        self.main_frame.grid(row=1, column=0, rowspan=3, columnspan=4, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_rowconfigure(15, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.project_label = customtkinter.CTkLabel(self.main_frame, text="MECHANIC HOURLY RATES", corner_radius=5, text_color='#f5f5f5', fg_color='#333', font=self.large_bold)
        self.project_label.grid(row=0, column=0, ipadx=10, ipady=1, padx=15, pady=(10,0), sticky='w')

        title_frame = customtkinter.CTkFrame(self.main_frame, corner_radius=0, fg_color='#f5f5f5', border_width=1, border_color='#777')
        title_frame.grid(row=1, column=0, columnspan=4, padx=0, pady=(10, 5), sticky="nsew")
        title_frame.grid_rowconfigure(15, weight=1)
        title_frame.grid_columnconfigure(0, weight=1)

        mechanic_label = customtkinter.CTkLabel(title_frame, text=f"NAME", font=self.bold_font)
        mechanic_label.grid(row=0, column=0, padx=25, pady=5, sticky='w')

        mechanic_label = customtkinter.CTkLabel(title_frame, text=f"HOURLY RATE ($ / Hr)", font=self.bold_font)
        mechanic_label.grid(row=0, column=1, padx=25, pady=5, sticky='w')

    def apply_settings(self, settings: dict):
        self.settings = settings
        try:
            mechanic_settings = self.settings['Mechanic Hourly Rates']
        except:
            mechanic_settings = {}
        
        self.entry_label_list = []
        self.mechanic_name_list = []
        i = 2
        for mechanic, hourly_rate in mechanic_settings.items():
            mechanic_label = customtkinter.CTkLabel(self.main_frame, text=f"{mechanic.upper()}", font=self.regular_font)
            mechanic_label.grid(row=i, column=0, padx=25, pady=5, sticky='w')

            rate_entry = customtkinter.CTkEntry(self.main_frame, width=140, font=self.regular_font)
            rate_entry.grid(row=i, column=1, padx=25, pady=3, sticky='')
            rate_entry.bind("<Return>", self.activate_save)
            try:
                rate_entry.insert(0, f"{hourly_rate}")
            except:
                rate_entry.insert(0, "110")
          
            self.mechanic_name_list.append(mechanic)
            self.entry_label_list.append(rate_entry)

            i += 1

        self.save_button = customtkinter.CTkButton(self.main_frame, width=40, text="SAVE", font=self.regular_font, command=self.activate_save)
        self.save_button.grid(row=i, column=0, ipadx=30, ipady=0, padx=25, pady=10, sticky='w')

    def activate_save(self, entry_field: str = ''):
        self.main_frame.focus()
        error = 0
        for i in range(len(self.mechanic_name_list)):
            name = self.mechanic_name_list[i]
            value_entry = self.entry_label_list[i].get().replace(',','')
            try:
                float_entry = round(float(value_entry or 0))
                self.settings['Mechanic Hourly Rates'].update({name:float_entry})
            except:
                error += 1
                self.entry_label_list[i].delete(0, len(value_entry))
            
        self.parent.activate_refresh()
        if error:
            messagebox.showerror('Number Input Error', 'Please Enter Number Only.')

