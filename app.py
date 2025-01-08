import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import pandas as pd
import json

# Global variables
start_x, start_y = None, None
current_rectangle = None
rectangles = {}
certificate_path = None


def generate_template():
    """Generate a sheet template based on user input."""
    try:
        num_entries = int(entries_input.get())
        num_fields = int(fields_input.get())

        # Create a DataFrame with specified number of entries and fields
        data = {f"Field_{i + 1}": ["" for _ in range(num_entries)] for i in range(num_fields)}
        df = pd.DataFrame(data)

        # Save as Excel file
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")],
                                                 title="Save Template File")
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "Template generated and saved successfully!")
    except ValueError:
        messagebox.showerror("Error", "Please enter valid numbers for entries and fields.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")


def upload_pdf():
    """Upload the certificate PDF."""
    global certificate_path
    certificate_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")], title="Select Certificate PDF")
    if certificate_path:
        messagebox.showinfo("Success", f"Certificate uploaded: {certificate_path}")


def start_rectangle(event):
    """Start drawing the rectangle on mouse click."""
    global start_x, start_y, current_rectangle
    start_x, start_y = event.x, event.y
    current_rectangle = canvas.create_rectangle(start_x, start_y, start_x, start_y, outline="red", width=2)


def update_rectangle(event):
    """Update the rectangle as the mouse is dragged."""
    global current_rectangle
    canvas.coords(current_rectangle, start_x, start_y, event.x, event.y)


def finalize_rectangle(event):
    """Finalize the rectangle and save the coordinates."""
    global current_rectangle
    field_name = field_entry.get()
    if field_name:
        # Get rectangle coordinates
        x1, y1, x2, y2 = canvas.coords(current_rectangle)
        rectangles[field_name] = (x1, y1, x2, y2)
        print(f"Field '{field_name}' set at {x1, y1, x2, y2}")
    current_rectangle = None  # Reset for the next rectangle


def save_positions():
    """Save all the field positions to a JSON file."""
    if not rectangles:
        messagebox.showwarning("No Positions", "No field positions to save.")
        return

    # Save field positions to a JSON file
    save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if save_path:
        with open(save_path, "w") as f:
            json.dump(rectangles, f)
        messagebox.showinfo("Success", "Field positions saved successfully!")


def load_positions():
    """Load the field positions from a JSON file."""
    global rectangles
    load_path = filedialog.askopenfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if load_path:
        with open(load_path, "r") as f:
            rectangles = json.load(f)
        messagebox.showinfo("Success", "Field positions loaded successfully!")


def back_to_main():
    """Back to the main screen."""
    root.quit()
    # Optionally, restart the main app (if in separate window), for now it quits.


# Create the Tkinter window for the main app
root = tk.Tk()
root.title("Certificate Generator")
root.geometry("400x300")

# Main screen widgets
tk.Label(root, text="Number of Entries:").pack(pady=5)
entries_input = tk.Entry(root)
entries_input.pack(pady=5)

tk.Label(root, text="Number of Fields:").pack(pady=5)
fields_input = tk.Entry(root)
fields_input.pack(pady=5)

# Buttons for actions
tk.Button(root, text="Generate Sheet Template", command=generate_template).pack(pady=10)
tk.Button(root, text="Upload Certificate PDF", command=upload_pdf).pack(pady=10)

def load_pdf_preview(pdf_path):
    """Load the first page of the PDF as an image."""
    pdf_document = fitz.open(pdf_path)
    page = pdf_document[0]  # First page
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def edit_positions_screen():
    """Set up the Tkinter window for editing positions."""
    global certificate_path, canvas, field_entry

    if not certificate_path:
        messagebox.showerror("Error", "Please upload a certificate PDF first.")
        return

    # Create the Tkinter window for the edit positions screen
    edit_window = tk.Toplevel(root)
    edit_window.title("Edit Field Positions")
    edit_window.geometry("800x600")

    canvas = tk.Canvas(edit_window, width=800, height=600)
    canvas.pack()

    # Load PDF preview
    pdf_image = load_pdf_preview(certificate_path)
    pdf_photo = ImageTk.PhotoImage(pdf_image)
    canvas.create_image(0, 0, anchor=tk.NW, image=pdf_photo)

    # Entry for field name
    field_label = tk.Label(edit_window, text="Field Name:")
    field_label.pack(pady=5)
    field_entry = tk.Entry(edit_window)
    field_entry.pack(pady=5)

    # Buttons for actions
    tk.Button(edit_window, text="Save Positions", command=save_positions).pack(pady=5)
    tk.Button(edit_window, text="Load Positions", command=load_positions).pack(pady=5)
    tk.Button(edit_window, text="Back to Main", command=back_to_main).pack(pady=5)

    # Bind mouse events to draw rectangles
    canvas.bind("<Button-1>", start_rectangle)
    canvas.bind("<B1-Motion>", update_rectangle)
    canvas.bind("<ButtonRelease-1>", finalize_rectangle)

    edit_window.mainloop()


# Add button to go to the edit positions screen
tk.Button(root, text="Edit Field Positions", command=edit_positions_screen).pack(pady=20)

root.mainloop()
