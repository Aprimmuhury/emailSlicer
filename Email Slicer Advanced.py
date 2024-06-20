import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Scrollbar, Canvas, PhotoImage
import pandas as pd
from pandas import ExcelWriter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def slice_email(email):
    if '@' in email:
        username, domain = email.split('@')
        extension = domain.split('.')[-1]
        return username, domain, extension
    else:
        return email, 'Invalid', 'Invalid'

def process_single_email():
    email = entry.get()
    username, domain, extension = slice_email(email)
    result_label.config(text=f"Username: {username}, Domain: {domain}, Extension: .{extension}")

def load_text():
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_path:
        with open(file_path, 'r') as file:
            lines = file.read().splitlines()
        global df  # Make DataFrame accessible globally for saving to Excel
        df = pd.DataFrame(lines, columns=['Email'])
        df[['Username', 'Domain']] = df['Email'].apply(lambda x: pd.Series(slice_email(x)[:2]))
        df['Extension'] = df['Domain'].apply(lambda x: x.split('.')[-1] if '.' in x else 'Invalid')
        display_data(df)
        plot_data(df)

def display_data(df):
    tree.delete(*tree.get_children())
    for row in df.itertuples():
        tree.insert("", 'end', values=(row.Email, row.Username, row.Domain, row.Extension))

def plot_data(df):
    extension_count = df['Extension'].value_counts()
    fig, ax = plt.subplots(figsize=(4, 3))  # Adjusted figure size for better GUI fit
    extension_count.plot(kind='bar', ax=ax)
    ax.set_title("Distribution of Email Extensions")
    ax.set_xlabel("Extensions")
    ax.set_ylabel("Counts")

    # Rotate the x-axis labels for better visibility
    plt.xticks(rotation=45, ha='right', fontsize=8)  # Rotate labels and adjust font size

    # Adjust the layout to make sure everything fits without clipping
    plt.tight_layout()

    # Update the canvas for the plot
    canvas = FigureCanvasTkAgg(fig, master=plot_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


def save_to_excel():
    if 'df' in globals():
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if filepath:
            df.to_excel(filepath, index=False)
            messagebox.showinfo("Information", "File saved successfully")
    else:
        messagebox.showerror("Error", "No data to save")

# Set up the main window
window = tk.Tk()
window.title('Simple Email Slicer')
window.geometry("1000x900")
window.resizable(width=False,height=False)
window.config(bg="#BE361A")

# Scrollable Canvas
canvas = Canvas(window)
scroll_y = Scrollbar(window, orient="vertical", command=canvas.yview)
frame = tk.Frame(canvas)  # Add a frame in the canvas

# Pack the scroll components
canvas.create_window((0, 0), window=frame, anchor='nw')
canvas.configure(yscrollcommand=scroll_y.set)
canvas.pack(fill='both', expand=True, side='left')
scroll_y.pack(fill='y', side='right')

# Entry widget for single email processing
entry = tk.Entry(frame, width=50)
entry.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

process_button = tk.Button(frame, text="Process Email", command=process_single_email)
process_button.grid(row=0, column=3, padx=10, pady=10)

result_label = tk.Label(frame, text="")
result_label.grid(row=1, column=0, columnspan=4)

load_button = tk.Button(frame, text="Load TXT", command=load_text)
load_button.grid(row=2, column=0, columnspan=4)

tree = ttk.Treeview(frame, columns=("Email", "Username", "Domain", "Extension"), show="headings")
for col in ("Email", "Username", "Domain", "Extension"):
    tree.heading(col, text=col)
tree.grid(row=3, column=0, columnspan=4, sticky='nsew')

save_button = tk.Button(frame, text="Save to Excel", command=save_to_excel)
save_button.grid(row=4, column=0, columnspan=4)

# Frame for plotting
plot_frame = tk.Frame(frame)
plot_frame.grid(row=5, column=0, columnspan=4, sticky='nsew')

# Update the scrollregion when all widgets are in canvas
frame.update_idletasks()
canvas.config(scrollregion=canvas.bbox("all"))

window.mainloop()
