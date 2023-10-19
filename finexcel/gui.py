import tkinter as tk
import style

# Action when button is cliked
def button1_click():
    style.generateSpreadsheet(entry1.get(), entry2.get())


# Creating the main window
root = tk.Tk()
root.title("FinExcel")

# Making the window fixed in size
root.geometry("600x150")
root.resizable(False, False)

# Creating the first input field
label1 = tk.Label(root, text="Ticker")
label1.pack()
entry1 = tk.Entry(root)
entry1.pack()

# Creating the second input field
label2 = tk.Label(root, text="API Key")
label2.pack()
entry2 = tk.Entry(root)
entry2.pack()


# Creating button to generate spreadsheet
button1 = tk.Button(root, text="Generate Spreadsheet", command=button1_click)
button1.pack()

# Creating a label to display the result
result_label = tk.Label(root, text="")
result_label.pack()


# Starting the tkinter main loop
root.mainloop()
