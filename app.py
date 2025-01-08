import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import json
import pandas as pd
import fitz  # PyMuPDF

# Global variables
rectangles = []  # List to store the coordinates of the rectangles
certificate_path = None  # Path to the uploaded certificate
data_path = None  # Path to the uploaded data sheet
field_names = []  # List to store field names from the data sheet


# Function to load the PDF preview
def load_pdf_preview(pdf_path):
    """Load the first page of the PDF as an image."""
    pdf_document = fitz.open(pdf_path)
    page = pdf_document[0]  # First page
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img


# Start drawing the rectangle
def start_rectangle(event):
    """Start drawing the rectangle."""
    global current_rectangle, start_x, start_y
    start_x = event.x
    start_y = event.y
    current_rectangle = canvas.create_rectangle(start_x, start_y, start_x, start_y, outline="red", width=2)


# Update the rectangle as user drags the mouse
def update_rectangle(event):
    """Update the rectangle as user drags the mouse."""
    canvas.coords(current_rectangle, start_x, start_y, event.x, event.y)


# Finalize the rectangle when the user releases the mouse button
def finalize_rectangle(event):
    """Finalize the rectangle and save the coordinates."""
    global current_rectangle
    if len(rectangles) < len(field_names):  # Only add if there are still fields to assign
        # Get rectangle coordinates
        x1, y1, x2, y2 = canvas.coords(current_rectangle)  # Coordinates of the rectangle
        rectangles.append((x1, y1, x2, y2))  # Save coordinates in list
        print(f"Field '{field_names[len(rectangles) - 1]}' set at {x1, y1, x2, y2}")
    current_rectangle = None  # Reset for the next rectangle


# Save all positions to a JSON file
def save_positions():
    """Save all the field positions to a JSON file."""
    if not rectangles:
        messagebox.showwarning("No Positions", "No field positions to save.")
        return

    # Save field positions to a JSON file
    save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if save_path:
        with open(save_path, "w") as f:
            json.dump(rectangles, f)  # Save the coordinates as JSON
        messagebox.showinfo("Success", "Field positions saved successfully!")


# Load field positions from a JSON file
def load_positions():
    """Load the field positions from a JSON file."""
    global rectangles
    load_path = filedialog.askopenfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if load_path:
        with open(load_path, "r") as f:
            rectangles = json.load(f)  # Load the coordinates from the JSON file
        messagebox.showinfo("Success", "Field positions loaded successfully!")


# Save positions and close the window after confirmation
def save_and_close():
    """Save positions and close the window after user confirmation."""
    confirmation = messagebox.askyesno("Confirm",
                                       "Are you sure all positions are correct? Do you want to save and close?")
    if confirmation:
        save_positions()
        edit_window.destroy()


# Go back to the main screen
def back_to_main():
    """Return to the main screen."""
    edit_window.destroy()


# Load the data sheet (Excel or CSV) for field values
def load_data_sheet():
    """Allow the user to upload the data sheet and load it."""
    global data_path, field_names
    data_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")])
    if data_path:
        # Load the data into a Pandas DataFrame
        if data_path.endswith(".xlsx"):
            df = pd.read_excel(data_path, engine="openpyxl")  # Specify the engine explicitly for .xlsx
        elif data_path.endswith(".xls"):
            df = pd.read_excel(data_path, engine="xlrd")  # Use xlrd for .xls files
        else:
            df = pd.read_csv(data_path)

        # Extract the column headers as field names
        field_names = list(df.columns)
        print(f"Loaded fields: {field_names}")
        messagebox.showinfo("Data Loaded", "Data sheet loaded successfully!")
        return df
    else:
        messagebox.showwarning("No File", "No data sheet selected.")
        return None


# The main screen for editing field positions
def edit_positions_screen(root):
    """Set up the Tkinter window for editing positions."""
    global certificate_path, canvas, edit_window

    if not certificate_path:
        messagebox.showerror("Error", "Please upload a certificate PDF first.")
        return

    if not field_names:
        messagebox.showerror("Error", "Please upload a data sheet first.")
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

    # Buttons for actions
    tk.Button(edit_window, text="Save Positions", command=save_positions).pack(pady=5)
    tk.Button(edit_window, text="Load Positions", command=load_positions).pack(pady=5)
    tk.Button(edit_window, text="Save and Close", command=save_and_close).pack(pady=5)
    tk.Button(edit_window, text="Back to Main", command=back_to_main).pack(pady=5)

    # Bind mouse events to draw rectangles
    canvas.bind("<Button-1>", start_rectangle)
    canvas.bind("<B1-Motion>", update_rectangle)
    canvas.bind("<ButtonRelease-1>", finalize_rectangle)

    edit_window.mainloop()


# Set up the main window
def main_screen():
    """Set up the main screen where the user uploads the certificate PDF and data sheet."""
    global certificate_path, data_path
    root = tk.Tk()
    root.title("Certificate Generator")
    root.geometry("400x400")

    def upload_certificate():
        """Allow the user to upload the certificate PDF."""
        global certificate_path
        certificate_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if certificate_path:
            messagebox.showinfo("Success", "Certificate PDF uploaded successfully!")

    def upload_data_sheet():
        """Allow the user to upload the data sheet."""
        load_data_sheet()

    tk.Label(root, text="Upload Certificate PDF:").pack(pady=10)
    tk.Button(root, text="Upload PDF", command=upload_certificate).pack(pady=10)

    tk.Label(root, text="Upload Data Sheet:").pack(pady=10)
    tk.Button(root, text="Upload Data Sheet", command=upload_data_sheet).pack(pady=10)

    tk.Button(root, text="Edit Field Positions", command=lambda: edit_positions_screen(root)).pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main_screen()
