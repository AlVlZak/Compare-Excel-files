# This script is combination of Find difference and Compare definition
__version__ = "1.1.0"
import tkinter as tk

import compare_excels as fd
import compare_by_key as cd


def setup_gui():
    # Creation of GUI

    def exit_():
        root.destroy()
        root.quit()

    root = tk.Tk()
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = 400
    window_height = 150
    center_x = int(screen_width / 2 - window_width / 2)
    center_y = int(screen_height / 2 - window_height / 2)

    root.title("Find difference")
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    btn1 = tk.Button(
        root,
        text='Compare two excels',
        command=lambda: fd.start(root),
        padx=15,
        pady=5
    )

    btn1.pack(expand=True, side=tk.TOP)

    btn2 = tk.Button(
        root,
        text='Compare to excels by unique key',
        command=lambda: cd.start(root),
        padx=20,
        pady=5
    )
    btn2.pack(expand=True, side=tk.TOP)

    btn3 = tk.Button(
        root,
        text='Exit',
        command=exit_,
        padx=25,
        pady=5
    )
    btn3.pack(expand=True, side=tk.BOTTOM)

    tk.mainloop()


if __name__ == '__main__':
    setup_gui()
