import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pandastable import Table
from sqlalchemy import create_engine

# Create the main window
root = tk.Tk()
root.title("ExcelSQL Studio")
wi_gui=900
hi_gui=670

wi_scr=root.winfo_screenwidth()
hi_scr=root.winfo_screenheight()

x=(wi_scr/2)-(wi_gui/2)
y=(hi_scr/2)-(hi_gui/2)

root.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))

# Global variables
excel_file = None
sheet_names = []
current_sheet_data = pd.DataFrame()
query_result = None
engine = create_engine("sqlite://", echo=False)

# Function to create a section with a header, footer, and content
def create_section(parent, header_text, footer_text, bg_color):
    section_frame = tk.Frame(parent, relief=tk.RAISED, borderwidth=2, bg=bg_color)

    # Header
    header = tk.Label(
        section_frame,
        text=header_text,
        bg="#2f714d",
        font=("Arial", 14, "bold"),
        height=2,
        anchor="w",
        padx=10,
        fg="white",
    )
    header.pack(fill=tk.X)

    # Content Frame
    content_frame = tk.Frame(section_frame, bg="white")
    content_frame.pack(fill=tk.BOTH, expand=True)

    # Footer
    footer = tk.Label(
        section_frame,
        text=footer_text,
        bg="lightgray",
        font=("Arial", 10, "italic"),
        height=1,
        anchor="w",
        padx=10,
    )
    footer.pack(fill=tk.X)

    return section_frame, content_frame, footer

# Callback for file upload
def upload_file(event=None):
    global excel_file, sheet_names, engine
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
    )
    if file_path:
        footer_file_upload.config(text=f"File: {file_path}")
        try:
            # Reset the engine
            engine = create_engine("sqlite://", echo=False)

            # Load Excel file and sheets
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            populate_sheet_list(sheet_names)

            # Load sheets into SQLite engine
            for sheet in sheet_names:
                sanitized_name = sheet.replace(" ", "_")
                df = excel_file.parse(sheet)
                df.to_sql(sanitized_name, engine, if_exists="replace", index=False)
            messagebox.showinfo("Success", "Sheets loaded into the database as tables.")
            select_sheet(sheet_names[0])  # Select the first sheet by default
        except Exception as e:
            footer_file_upload.config(text=f"Error: {str(e)}")

# Populate the sheet names in the listbox
def populate_sheet_list(sheets):
    sheet_listbox.delete(0, tk.END)
    for sheet in sheets:
        sheet_listbox.insert(tk.END, sheet)

# Handle sheet selection
def select_sheet(sheet_name):
    global current_sheet_data
    sanitized_name = sheet_name.replace(" ", "_")
    try:
        current_sheet_data = pd.read_sql_query(f"SELECT * FROM {sanitized_name}", engine)
        total_rows, total_columns = current_sheet_data.shape
        footer_sheet_names.config(
            text=f"{sheet_name}: {total_rows} X {total_columns}"
        )

        # footer_sheet_names.config(
        #     text=f"Sheet: {sheet_name}, Columns: {total_columns}, Rows: {total_rows}"
        # )
    except Exception as e:
        footer_sheet_names.config(text=f"Error: {str(e)}")

# Execute SQL Query
def execute_query():
    global query_result
    query = query_textbox.get("1.0", tk.END).strip()
    if not query:
        messagebox.showerror("Error", "Query cannot be empty.")
        return
    try:
        query_result = pd.read_sql_query(query, engine)
        update_result_table(query_result)
        total_rows, total_columns = query_result.shape
        footer_result.config(
            text=f"Columns: {total_columns}, Rows: {total_rows}"
        )
        messagebox.showinfo("Success", "Query executed successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))



# Export result
def export_result(file_type):
    if query_result is None or query_result.empty:
        messagebox.showerror("Error", "No result to export.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=f".{file_type}", filetypes=[(file_type.upper(), f"*.{file_type}")])
    if not file_path:
        return
    try:
        if file_type == "xlsx":
            query_result.to_excel(file_path, index=False)
        elif file_type == "csv":
            query_result.to_csv(file_path, index=False)
        elif file_type == "txt":
            query_result.to_csv(file_path, index=False, sep='\t')
        messagebox.showinfo("Success", f"Result exported to {file_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# On listbox selection change
def on_sheet_select(event):
    selected_index = sheet_listbox.curselection()
    if selected_index:
        selected_sheet = sheet_listbox.get(selected_index)
        select_sheet(selected_sheet)

# Main PanedWindow
main_paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
main_paned.pack(fill=tk.BOTH, expand=True)

# Left Pane
left_pane = ttk.PanedWindow(main_paned, orient=tk.VERTICAL)
main_paned.add(left_pane, weight=1)

# File Upload Section
file_upload_frame, file_upload_content, footer_file_upload = create_section(
    left_pane, header_text="File Upload", footer_text="No file uploaded.", bg_color="white"
)
left_pane.add(file_upload_frame, weight=1)

# Drag-and-Drop and Clickable Upload Setup
drag_label = tk.Label(
    file_upload_content,
    text="Drag and drop an Excel file here\nor click below to upload",
    bg="white",
    font=("Arial", 10),
    pady=20,
)
drag_label.pack(fill=tk.BOTH, expand=True)

upload_button = tk.Button(
    file_upload_content,
    text="Click to Upload File",
    command=upload_file,
    font=("Arial", 10, "bold"),
    bg="green",
    fg="white",
)
upload_button.pack(pady=10)

# Sheet Names Section
sheet_names_frame, sheet_names_content, footer_sheet_names = create_section(
    left_pane, header_text="Sheet Names", footer_text="Select a sheet", bg_color="white"
)
left_pane.add(sheet_names_frame, weight=1)

# Listbox with Scrollbar for sheet names
scrollbar = tk.Scrollbar(sheet_names_content)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

sheet_listbox = tk.Listbox(sheet_names_content, yscrollcommand=scrollbar.set)
sheet_listbox.pack(fill=tk.BOTH, expand=True)

scrollbar.config(command=sheet_listbox.yview)
sheet_listbox.bind("<<ListboxSelect>>", on_sheet_select)

# # Summary Section
# summary_frame, summary_content, footer_summary = create_section(
#     left_pane, header_text="Summary", footer_text="Sheet info will appear here", bg_color="white"
# )
# left_pane.add(summary_frame, weight=1)

# Right Pane (Query and Result Sections)
right_pane = ttk.PanedWindow(main_paned, orient=tk.VERTICAL)
main_paned.add(right_pane, weight=2)

# Query Section
query_frame, query_content, footer_query = create_section(
    right_pane, header_text="Query Here", footer_text="", bg_color="white"
)
right_pane.add(query_frame, weight=1)

query_textbox = tk.Text(query_content, height=5, font=("Arial", 10))
query_textbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

run_query_button = tk.Button(
    query_content,
    text="Run Query",
    command=execute_query,
    font=("Arial", 10, "bold"),
    bg="green",
    fg="white",
)
run_query_button.pack(anchor="se", padx=10, pady=10)




# Result Section
result_frame, result_content, footer_result = create_section(
    right_pane, header_text="Result", footer_text="", bg_color="white"
)
right_pane.add(result_frame, weight=2)

# Create a canvas for scrolling
result_canvas = tk.Canvas(result_content, bg="white")
result_canvas.pack(fill=tk.BOTH, expand=False)

def update_result_table(data):
    # Clear previous content
    for widget in result_canvas.winfo_children():
        widget.destroy()

    # Create the PandasTable directly in the result content frame
    pt = Table(
        result_canvas,  # Attach PandasTable to the frame
        dataframe=data,
        showtoolbar=False,  # Disable PandasTable toolbar
        showstatusbar=False  # Disable PandasTable status bar
    )
    pt.show()

    # # Ensure PandasTable uses its own scrollbars and fits the frame
    # pt.redraw()





# Footer with Export Buttons (always visible)
export_button_frame = tk.Frame(result_frame, bg="lightgray")
export_button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

export_excel_button = tk.Button(
    export_button_frame,
    text="Export to Excel",
    command=lambda: export_result("xlsx"),
    bg="green",
    fg="white",
    font=("Arial", 10)
)
export_excel_button.pack(side=tk.RIGHT, padx=5)

export_csv_button = tk.Button(
    export_button_frame,
    text="Export to CSV",
    command=lambda: export_result("csv"),
    bg="green",
    fg="white",
    font=("Arial", 10)
)
export_csv_button.pack(side=tk.RIGHT, padx=5)

export_txt_button = tk.Button(
    export_button_frame,
    text="Export to TXT",
    command=lambda: export_result("txt"),
    bg="green",
    fg="white",
    font=("Arial", 10)
)
export_txt_button.pack(side=tk.RIGHT, padx=5)








# Run the Tkinter main loop
root.mainloop()
