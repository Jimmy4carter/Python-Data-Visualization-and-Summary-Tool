# Developer: Jimmy Carter
# Contact: jimmy4carter@gmail.com
# Phone: +2348038660259
# GitHub: https://github.com/Jimmy4carter


import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# Global variables to store the loaded data and summary
loaded_data = None
summary_stats = None

# Function to load data and create visualizations
def load_data():
    global loaded_data, summary_stats

    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            if file_path.endswith('.json'):
                loaded_data = pd.read_json(file_path)
            elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                loaded_data = pd.read_excel(file_path)
            else:
                raise ValueError("Invalid file format.")

            # Clear previous visualizations and summary
            clear_content()

            # Display heading with file name
            file_name = os.path.basename(file_path)
            heading_text = f"File: {file_name}\n\n"
            summary_text.insert(tk.END, heading_text)

            # Analysis: Unique Counts and Null Counts
            unique_counts = loaded_data.nunique()
            null_counts = loaded_data.isnull().sum()

            analysis_text = "Analysis:\n"
            analysis_text += "-----------------------------------------\n"
            analysis_text += "Unique Counts:\n"
            analysis_text += f"{unique_counts.to_string()}\n\n"
            analysis_text += "Null Counts:\n"
            analysis_text += f"{null_counts.to_string()}\n\n"

            summary_text.insert(tk.END, analysis_text)

            # Display summary statistics
            summary_stats = loaded_data.describe()
            summary_text.insert(tk.END, f"Summary Statistics:\n{summary_stats.to_string()}\n\n")

            # Create visualizations
            numeric_columns = loaded_data.select_dtypes(include='number').columns.tolist()
            categorical_columns = loaded_data.select_dtypes(exclude='number').columns.tolist()

            if numeric_columns:
                for col in numeric_columns:
                    plt.figure(figsize=(6, 4))
                    plt.hist(loaded_data[col], bins=10)
                    plt.title(f'Histogram of {col}')
                    plt.xlabel('Values')
                    plt.ylabel('Frequency')

                    # Display the histogram on Tkinter window
                    canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)
                    canvas.draw()
                    canvas.get_tk_widget().pack()

            if categorical_columns:
                for col in categorical_columns:
                    value_counts = loaded_data[col].value_counts()
                    plt.figure(figsize=(8, 6))
                    value_counts.plot(kind='bar')
                    plt.title(f'Bar Chart of {col}')
                    plt.xlabel('Categories')
                    plt.ylabel('Count')

                    # Display the bar chart on Tkinter window
                    canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)
                    canvas.draw()
                    canvas.get_tk_widget().pack()

                # Pie chart for the first categorical column
                first_categorical_column = categorical_columns[0]
                plt.figure(figsize=(6, 6))
                loaded_data[first_categorical_column].value_counts().plot(kind='pie', autopct='%1.1f%%')
                plt.title(f'Pie Chart of {first_categorical_column}')
                plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle

                # Display the pie chart on Tkinter window
                canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)
                canvas.draw()
                canvas.get_tk_widget().pack()

            # Additional visualizations
            if len(numeric_columns) > 1:
                plt.figure(figsize=(8, 6))
                loaded_data[numeric_columns[:2]].plot.scatter(x=numeric_columns[0], y=numeric_columns[1])
                plt.title(f'Scatter Plot of {numeric_columns[0]} vs {numeric_columns[1]}')
                plt.xlabel(numeric_columns[0])
                plt.ylabel(numeric_columns[1])

                # Display the scatter plot on Tkinter window
                canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)
                canvas.draw()
                canvas.get_tk_widget().pack()

        except Exception as e:
            messagebox.showerror("Error", str(e))

# Function to clear content in the summary and visualization frames
def clear_content():
    summary_text.delete('1.0', tk.END)
    for widget in plot_frame.winfo_children():
        widget.destroy()

# Function to export report to Excel
def export_to_excel():
    if loaded_data is not None and summary_stats is not None:
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    loaded_data.to_excel(writer, sheet_name='Original Data')
                    summary_stats.to_excel(writer, sheet_name='Summary Statistics')

                    # Export summary_text content to a separate sheet
                    summary_text_df = pd.DataFrame([summary_text.get("1.0", tk.END)], columns=["Summary Text"])
                    summary_text_df.to_excel(writer, sheet_name='Summary Text', index=False)

                messagebox.showinfo("Exported", "Summary report exported to Excel successfully!")

        except Exception as e:
            messagebox.showerror("Export Error", str(e))
    else:
        messagebox.showwarning("Export Warning", "Load data and generate a report before exporting!")


# Create the main window
root = tk.Tk()
root.title("Data Visualization and Summary")
root.geometry("1000x800")

# Create a menu bar
menubar = tk.Menu(root)
root.config(menu=menubar)

file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Load File", command=load_data)
file_menu.add_separator()
file_menu.add_command(label="Export to Excel", command=export_to_excel)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.destroy)
menubar.add_cascade(label="File", menu=file_menu)

# Create a frame for content
content_frame = tk.Frame(root)
content_frame.pack(fill=tk.BOTH, expand=True)

# Frame for summary
summary_frame = tk.Frame(content_frame, bd=2, relief=tk.GROOVE)
summary_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

summary_label = tk.Label(summary_frame, text="Summary", font=("Arial", 14, "bold"))
summary_label.pack()

# Create a scrollable text box for summary
summary_text = tk.Text(summary_frame, height=30, width=40, font=("Arial", 10))
summary_text.pack(fill=tk.BOTH, expand=True)

# Create a scrollbar for the summary text box
summary_scrollbar = tk.Scrollbar(summary_frame, command=summary_text.yview)
summary_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
summary_text.config(yscrollcommand=summary_scrollbar.set)

# Create a frame for scrollable visualizations
plot_frame = tk.Frame(content_frame, bd=2, relief=tk.GROOVE)
plot_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

root.mainloop()
