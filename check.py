import tkinter as tk
from tkinter import ttk, filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import pandas as pd  # For handling Excel files

# Create the main window
root = TkinterDnD.Tk()
root.title("Excel File Viewer")
root.geometry("800x600")

# Function to create a section with a header, footer, and content
def create_section(parent, header_text, footer_text, bg_color):
    section_frame = tk.Frame(parent, relief=tk.RAISED, borderwidth=2, bg=bg_color)

    # Header
    header = tk.Label(
        section_frame,
        text=header_text,
        bg="lightblue",
        font=("Arial", 12, "bold"),
        height=2,
        anchor="w",
        padx=10,
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
    global excel_file, sheet_names
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
    )
    if file_path:
        footer_file_upload.config(text=f"File: {file_path}")
        try:
            excel_file = pd.ExcelFile(file_path)  # Load the Excel file
            sheet_names = excel_file.sheet_names  # Get sheet names
            populate_sheet_list(sheet_names)  # Populate the sheet names list
            select_sheet(sheet_names[0])  # Select the active sheet by default
        except Exception as e:
            footer_file_upload.config(text=f"Error: {str(e)}")

# Callback for drag-and-drop
def drop_file(event):
    global excel_file, sheet_names
    file_path = event.data.strip('{').strip('}')  # Extract file path from event
    if file_path.endswith((".xls", ".xlsx")):  # Ensure it's an Excel file
        footer_file_upload.config(text=f"File: {file_path}")
        try:
            excel_file = pd.ExcelFile(file_path)  # Load the Excel file
            sheet_names = excel_file.sheet_names  # Get sheet names
            populate_sheet_list(sheet_names)  # Populate the sheet names list
            select_sheet(sheet_names[0])  # Select the active sheet by default
        except Exception as e:
            footer_file_upload.config(text=f"Error: {str(e)}")
    else:
        footer_file_upload.config(text="Error: Please upload a valid Excel file.")


# Populate the sheet names in the listbox
def populate_sheet_list(sheets):
    sheet_listbox.delete(0, tk.END)  # Clear the listbox
    for sheet in sheets:
        sheet_listbox.insert(tk.END, sheet)  # Add sheet names to the listbox

# Handle sheet selection
def select_sheet(sheet_name):
    global excel_file
    try:
        sheet_data = excel_file.parse(sheet_name)  # Parse the selected sheet
        total_rows = sheet_data.shape[0]
        total_columns = sheet_data.shape[1]
        footer_summary.config(
            text=f"Sheet: {sheet_name}, Columns: {total_columns}, Rows: {total_rows}"
        )
    except Exception as e:
        footer_summary.config(text=f"Error: {str(e)}")

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

drag_label.drop_target_register(DND_FILES)
drag_label.dnd_bind("<<Drop>>", drop_file)

upload_button = tk.Button(
    file_upload_content,
    text="Click to Upload File",
    command=upload_file,
    font=("Arial", 10),
    bg="lightblue",
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

# Summary Section
summary_frame, summary_content, footer_summary = create_section(
    left_pane, header_text="Summary", footer_text="Sheet info will appear here", bg_color="white"
)
left_pane.add(summary_frame, weight=1)

# Right Pane (Query and Result Sections)
right_pane = ttk.PanedWindow(main_paned, orient=tk.VERTICAL)
main_paned.add(right_pane, weight=2)

# Query Section
query_frame = create_section(
    right_pane, header_text="Query Here", footer_text="Footer: Query", bg_color="white"
)[0]
right_pane.add(query_frame, weight=1)

# Result Section
result_frame = create_section(
    right_pane, header_text="Result", footer_text="Footer: Result", bg_color="white"
)[0]
right_pane.add(result_frame, weight=2)

# Main Footer
main_footer = tk.Frame(root, relief=tk.RAISED, borderwidth=2, bg="lightcoral")
main_footer.pack(fill=tk.X)
footer_label = tk.Label(main_footer, text="Main Footer", bg="lightcoral", font=("Arial", 12, "bold"))
footer_label.pack()