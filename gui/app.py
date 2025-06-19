import tkinter as tk
from gui.file_select_frame import FileSelectFrame
from gui.mapping_frame import MappingFrame

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("データ移行ツール")
        self.geometry("700x500")
        self.active_frame = None
        self.launch_file_select()

    def launch_file_select(self):
        if self.active_frame:
            self.active_frame.destroy()
        self.active_frame = FileSelectFrame(self, self.launch_mapping_screen)
        self.active_frame.pack(fill="both", expand=True)

    def launch_mapping_screen(self, src_df, src_cols, dst_cols):
        if self.active_frame:
            self.active_frame.destroy()
        self.active_frame = MappingFrame(self, src_df, src_cols, dst_cols, self.launch_file_select)
        self.active_frame.pack(fill="both", expand=True)

def launch_app():
    app = App()
    app.mainloop()