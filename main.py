# main.py
from species_processor import SpeciesProcessorApp
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = SpeciesProcessorApp(root)
    root.mainloop()