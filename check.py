import tkinter as tk
import re

# Placeholder text setup
placeholder_text = "Click here to enter your query...\nFor Faster Load use CSV file with short File or sheet name\nNote: If a sheet name or column name contains spaces, replace them with underscores ('_') in query."

# Function to handle focus in event (clear placeholder text)
def on_focus_in(event):
    if query_textbox.get("1.0", "end-1c") == placeholder_text:
        query_textbox.delete("1.0", "end")
        query_textbox.config(fg="black")

# Function to handle focus out event (restore placeholder text if text is empty)
def on_focus_out(event):
    if query_textbox.get("1.0", "end-1c") == "":
        query_textbox.insert("1.0", placeholder_text)
        query_textbox.config(fg="gray")

# SQLAlchemy keywords and functions for highlighting
sqlalchemy_keywords = [
    "select", "insert", "update", "delete", "from", "where", "and", "or", 
    "like", "join", "inner", "left", "right", "outer", "limit", "group", "order", 
    "asc", "desc", "func", "distinct", "having", "between", "in", "is", "null", 
    "ilike", "exists", "not", "on", "as", "alias", "case", "when", "then",
    "else", "end", "all", "any", "some", "exists", "with", "having", "union", "intersect", 
    "except", "using", "corresponding", "left outer", "right outer", "full outer",
    "count", "sum", "avg", "min", "max", "coalesce", "concat", "round", "length", "lower",
    "upper", "now", "current_date", "current_time", "date", "date_trunc", "extract", "day", 
    "month", "year", "hour", "minute", "second", "random", "position", "mod", "concat_ws",
    "like", "ilike", "trim", "replace", "regexp_match", "regexp_replace", "cast", "isnull", 
    "isnotnull", "ifnull", "date_part", "jsonb_extract", "json_extract", "json_each", "json_array_length",
    "json_populate_record", "jsonb_array_elements_text", "jsonb_set", "jsonb_object_keys", "current_user",
    "current_schema", "current_catalog"
]

# Function to highlight SQLAlchemy keywords and functions in bold and color
def highlight_keywords(event=None):
    user_input = query_textbox.get("1.0", "end-1c")
    
    # Clear previous tags (if any)
    query_textbox.tag_remove("keyword", "1.0", "end")

    # Match keywords or functions (consider functions with parentheses)
    words = re.findall(r'\b(?:' + '|'.join(re.escape(keyword) for keyword in sqlalchemy_keywords) + r')\b(?:\s*\(.*?\))?', user_input)

    idx = "1.0"
    for word in words:
        # Find all occurrences of each word in the input and highlight if it matches a keyword
        start_idx = query_textbox.search(word, idx, stopindex=tk.END)
        while start_idx:
            end_idx = f"{start_idx}+{len(word)}c"
            query_textbox.tag_add("keyword", start_idx, end_idx)
            idx = end_idx  # Update index to continue search after the current match
            start_idx = query_textbox.search(word, idx, stopindex=tk.END)

    # Configure the keyword tag to make the text bold and colored
    query_textbox.tag_configure("keyword", foreground="blue", font=("Arial", 12, "bold"))

# Create the query textbox with placeholder
root = tk.Tk()
query_textbox = tk.Text(root, height=5, font=("Arial", 11), fg="gray")
query_textbox.insert("1.0", placeholder_text)
query_textbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# Bind events for focus in and out
query_textbox.bind("<FocusIn>", on_focus_in)
query_textbox.bind("<FocusOut>", on_focus_out)

# Key binding to trigger keyword highlighting on typing
query_textbox.bind("<KeyRelease>", highlight_keywords)

# Run the main loop
root.mainloop() 