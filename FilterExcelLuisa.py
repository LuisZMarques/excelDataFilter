import tkinter
import tkinter.messagebox
import customtkinter
import openpyxl

customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Filtro de Excel - Luisa")
        self.geometry(f"{1280}x{768}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # Sidebar

        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        # Sidebar content

        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="ExceLuisa", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

    
        # Main Frame
        
        self.main_frame = customtkinter.CTkFrame(self, width=500, corner_radius=10)
        self.main_frame.grid(row=0, column=1, rowspan=3, padx=(20, 20), pady=(100, 100), sticky="nsew")


        # Checkbox Frame

        self.checkbox_frame = customtkinter.CTkFrame(self, width=100, corner_radius=10)
        self.checkbox_frame.grid(row=0, column=2, rowspan=3, padx=(20, 20), pady=(100, 100), sticky="nsew")

        # Main Frame content
        
        self.button_1 = customtkinter.CTkButton(self.main_frame, command=self.button_openFile, text="Abrir Ficheiro")
        self.button_1.grid(row=1, column=0, padx=(20, 20), pady=(20, 20))
        self.button_2 = customtkinter.CTkButton(self.main_frame, command=self.button_closeFile, text="Guardar Ficheiro")
        self.button_2.grid(row=6, column=0, padx=(20, 20),pady=(20, 20))

        # Checkbox Frame

        self.checkbox_frame = customtkinter.CTkFrame(self, width=100, corner_radius=10)
        self.checkbox_frame.grid(row=0, column=2, rowspan=3, padx=(20, 20), pady=(100, 100), sticky="nsew")


        # set default values
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def button_openFile(self):
        print("teste open")
    
    def button_closeFile(self):
        print("teste close")



if __name__ == "__main__":
    app = App()
    app.mainloop()
