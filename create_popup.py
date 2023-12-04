import tkinter as tk
from tkinter import ttk
from tkinter import font
from tkinter.messagebox import askquestion

def popup_showinfo(item_name):
    root = tk.Tk()
    my_var = askquestion("Window", f"{item_name} já existe neste grupo!\nAdicionar novo valor ao valor já existente ou substituir?\nSim - Adicionar\nNão - Substituir", icon='warning')
    root.destroy()
    root.quit()
    if my_var == 'yes':
        return 1
    return 2

def popup_for_missing_requirements(requirement, filename):
    root = tk.Tk()
    root.title(f'{requirement} missing')
    root.geometry('1000x300')
    root.resizable(False, False)

    # INFO
    error_message = f'Requirement {requirement} is missing in {filename}!\nPlease open {filename} and indicate value that correspond to {requirement} down below!'
    error_message = 'Missing Requirement: '
    file_error = 'In File: '
    error_label = tk.Label(root, text=error_message)
    file_label = tk.Label(root, text=file_error)
    req = tk.Label(root, text=requirement, font=("Helvetica", 12, "bold"))
    file = tk.Label(root, text=filename, font=("Helvetica", 12, "bold"))

    # PACKING
    error_label.pack()
    req.pack()
    file_label.pack()
    file.pack()

    info = f'Please check file and insert the value that corresponds to "{requirement}" in the box below!'
    info_label = tk.Label(root, text=info)

    info_label.pack(pady=(30, 10))

    # ENTRY BOX
    value_var = tk.StringVar()
    entry = tk.Entry(root, textvariable=value_var)
    entry.pack()

    # FUNCS
    def get_value():
        root.destroy()
        root.quit()

    # CONTINUE BUTTON
    btn = tk.Button(root, text='CONTINUE', command=get_value)
    btn.pack(pady=10)

    root.mainloop()
    val = value_var.get()
    if val == '':
        return None
    return val
