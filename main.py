# -------------------------------------------------------------------------------------------------------------------- #
#                                                        IMPORTS                                                       #
# -------------------------------------------------------------------------------------------------------------------- #

#standard library imports
import os
from datetime import datetime

#third party imports
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document

#local imports
from tooltip import CreateToolTip

#changing the working directory, so that the program can be run from any path
os.chdir(os.path.dirname(os.path.realpath(__file__)))

# -------------------------------------------------------------------------------------------------------------------- #
#                                                 WINDOW INITIALIZATION                                                #
# -------------------------------------------------------------------------------------------------------------------- #

#window creation
window = tk.Tk()
window.title("Musikredaktion Rotationsmanager")
window.minsize(411, 378)

#tkinter variables
kalenderwoche = tk.StringVar(window, datetime.now().isocalendar()[1])

#keyboard bindings
def focus_next_widget(event):
	event.widget.tk_focusNext().focus()
	return("break")
def focus_prev_widget(event):
	event.widget.tk_focusPrev().focus()
	return("break")
window.bind("<Tab>", focus_next_widget)
window.bind("<Shift-Tab>", focus_prev_widget)

#theme: https://github.com/rdbende/Azure-ttk-theme
window.tk.call("source", os.path.join("theme", "azure.tcl"))
#window.tk.call("set_theme", "light")
window.tk.call("set_theme", "dark")
#ttk.Style().theme_use("azure")

def tkinter_center(win: tk.Tk | tk.Toplevel):
	"""Centers a tkinter window on the screen.
	Copied from https://stackoverflow.com/a/10018670"""
	win.update_idletasks()
	width = win.winfo_width()
	frm_width = win.winfo_rootx() - win.winfo_x()
	win_width = width + 2 * frm_width
	height = win.winfo_height()
	titlebar_height = win.winfo_rooty() - win.winfo_y()
	win_height = height + titlebar_height + frm_width
	x = win.winfo_screenwidth() // 2 - win_width // 2
	y = win.winfo_screenheight() // 2 - win_height // 2
	win.geometry(f"{width}x{height}+{x}+{y}")
	win.deiconify()

# -------------------------------------------------------------------------------------------------------------------- #
#                                                       COMMANDS                                                       #
# -------------------------------------------------------------------------------------------------------------------- #

filetypes = [("Textdatei", "*.txt"), ("Alle Dateien", "*.*")]

def import_file():
	"""Imports the text from a .txt file when the user clicks the import button."""
	try:
		file = filedialog.askopenfile("r", filetypes = filetypes, title = "Textdatei importieren")
		if file is not None:
			text.delete(1.0, tk.END)
			text.insert(tk.END, "".join(file.readlines()))
	except Exception as e:
		messagebox.showerror("Fehler beim Textdatei importieren", e)
	finally:
		if file and file is not None and not file.closed:
			file.close()

def export_file():
	"""Exports the text to a .txt file when the user clicks the export button."""
	try:
		file = filedialog.asksaveasfile("w", confirmoverwrite = True, defaultextension = ".txt", filetypes = filetypes, initialfile = "KW" + kalenderwoche.get(), title = "Textdatei exportieren")
		if file is not None:
			file.write(text.get("1.0", tk.END))
	except Exception as e:
		messagebox.showerror("Fehler beim Textdatei exportieren", e)
	finally:
		if file and file is not None and not file.closed:
			file.close()

def get_artist_and_title(string: str) -> str:
	index = string.find("--")
	return string if index == -1 else string[:index-1]

def generate_file():
	"""Generates a .docx file from the text input when the user clicks the corresponding button."""
	try:
		#open template file and check if it contains at least one table
		document = Document("template.docx")
		if len(document.tables) == 0:
			messagebox.showerror("Error", "Template.docx muss mindestens eine Tabelle enthalten.")
			return
		#replace the calendar week in the document
		for p in document.paragraphs:
			if "Playlisten-Rotation: KW" in p.text:
				for run in p.runs:
					if "KW" in run.text:
						run.text = run.text.replace("KW", "KW " + kalenderwoche.get())
						break
				break
		#create generator for the processed strings from the text widget
		line_generator = (string.strip() for string in text.get("1.0", tk.END).split("\n"))
		line_generator = map(get_artist_and_title, filter(None, line_generator))
		table = document.tables[0]
		#fill the table and dynamically add rows
		#ressource: https://python-docx.readthedocs.io/en/latest/api/table.html
		table.cell(0, 0).text = next(line_generator, "")
		for line in line_generator:
			table.add_row().cells[0].text = line
		#save the document
		filename = filedialog.asksaveasfilename(confirmoverwrite = True, defaultextension = ".docx", filetypes = [("Worddatei", "*.docx"), ("Alle Dateien", "*.*")], initialfile = "KW" + kalenderwoche.get(), title = "Worddatei exportieren")
		if filename == "":
			return
		if os.path.isdir(os.path.dirname(filename)):
			document.save(filename)
			#messagebox.showinfo("DOCX generieren", "DOCX wurde erfolgreich generiert.")
		else:
			messagebox.showerror("PathError", "Der Pfad ist ungültig.")
	except Exception as e:
		messagebox.showerror("Fehler beim Worddatei exportieren", e)

# -------------------------------------------------------------------------------------------------------------------- #
#                                                        LAYOUT                                                        #
# -------------------------------------------------------------------------------------------------------------------- #

padding = 16

#text input with scrollbar
frame_text = tk.Frame(window)
scrollbar = ttk.Scrollbar(frame_text)
text = tk.Text(frame_text, width = 50, height = 15, undo = True, yscrollcommand = scrollbar.set)
scrollbar.config(command = text.yview)
scrollbar.pack(side = tk.RIGHT, fill = tk.Y)
text.pack(fill = tk.BOTH, expand = True)
frame_text.pack(fill = tk.BOTH, expand = True)

frame_controls = tk.Frame(window)

#row with buttons
frame_first_row = tk.Frame(frame_controls)
button = ttk.Button(frame_first_row, text = "Importieren ⭳", cursor = "hand2", command = import_file)
CreateToolTip(button, "Importiert den Text aus einer Textdatei.")
button.pack(side = tk.LEFT)
button = ttk.Button(frame_first_row, text = "Exportieren ↥", cursor = "hand2", command = export_file)
CreateToolTip(button, "Exportiert den Text in eine Textdatei.")
button.pack(side = tk.LEFT, padx = padding)
button = ttk.Button(frame_first_row, text = "DOCX generieren →", cursor = "hand2", command = generate_file, style = "Accent.TButton")
CreateToolTip(button, "Generiert die Worddatei aus dem eingegebenen Text.")
button.pack()
frame_first_row.pack()

#row with labels and calendar week control
frame_second_row = tk.Frame(frame_controls)
ttk.Label(frame_second_row, text = "Kalenderwoche:").pack(side = tk.LEFT, padx = (0, 5))
ttk.Spinbox(frame_second_row, from_ = 1, to = 52, textvariable = kalenderwoche, width = 5).pack(side = tk.LEFT)
ttk.Label(frame_second_row, text = "© 2024 Jona Heinke\nunter MIT Lizenz").pack(side = tk.LEFT, padx = padding)
ttk.Label(frame_second_row, text = "Version: 1").pack(side = tk.LEFT)
frame_second_row.pack(pady = (padding, 0))

frame_controls.pack(side = tk.BOTTOM, fill = tk.X, padx = padding, pady = padding)

# -------------------------------------------------------------------------------------------------------------------- #
#                                                       MAIN LOOP                                                      #
# -------------------------------------------------------------------------------------------------------------------- #

tkinter_center(window)
window.mainloop()