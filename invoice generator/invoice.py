# Import necessary modules
import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

# Function to clear the input fields
def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

# List to store invoice items
invoice_list = []

# Function to add an item to the invoice
def add_item():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, desc, price, line_total]
    tree.insert('', 0, values=invoice_item)  # Insert item into the tree view
    clear_item()  # Clear input fields
    invoice_list.append(invoice_item)  # Add item to the invoice list

# Variables for editing mode
editing = False
editing_item_index = None

# Function to toggle between editing and saving an item
def toggle_edit_save():
    global editing
    selected_item = tree.selection()

    if not editing and selected_item:
        # If not editing and an item is selected, switch to editing mode
        editing = True
        edit_save_button.config(text="Save Item")
        add_item_button.config(state=tkinter.DISABLED)  # Disable other buttons while editing
        delete_item_button.config(state=tkinter.DISABLED)

        # Retrieve the selected item's ID (index)
        item_id = selected_item[0]
        # Store the index of the selected item for editing
        global editing_item_index
        editing_item_index = tree.index(item_id)

        # Retrieve the selected item's values and populate the input boxes
        item_values = tree.item(item_id, 'values')
        qty_spinbox.delete(0, tkinter.END)
        qty_spinbox.insert(0, item_values[0])
        desc_entry.delete(0, tkinter.END)
        desc_entry.insert(0, item_values[1])
        price_spinbox.delete(0, tkinter.END)
        price_spinbox.insert(0, item_values[2])
    elif editing:
        # If in editing mode, save the edited item and switch back to "Edit Item" mode
        editing = False
        edit_save_button.config(text="Edit Item")
        if editing_item_index is not None:
            # Retrieve the edited values directly from the input fields
            qty = int(qty_spinbox.get())
            desc = desc_entry.get()
            price = float(price_spinbox.get())
            line_total = qty * price

            # Update the item in the tree widget with the edited values
            tree.item(selected_item, values=(qty, desc, price, line_total))

            # Update the corresponding item in the invoice_list with the edited values
            invoice_list[editing_item_index] = [qty, desc, price, line_total]

            # Enable other buttons after saving
            add_item_button.config(state=tkinter.NORMAL)
            delete_item_button.config(state=tkinter.NORMAL)
            clear_item()  # Clear the input fields
            editing_item_index = None  # Reset the editing_item_index

# Function to delete an item
def delete_item():
    selected_item = tree.selection()
    if selected_item:
        tree.delete(selected_item)  # Delete the selected item from the tree view

# Function to start a new invoice
def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()  # Clear input fields
    tree.delete(*tree.get_children())  # Clear the tree view
    invoice_list.clear()  # Clear the invoice list

# Function to generate the invoice
def generate_invoice():
    doc = DocxTemplate("invoice generator/invoice_template.docx")
    name = first_name_entry.get() + last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.15
    total = subtotal + (subtotal * salestax)
    total = round(total, 2)

    # Pass the updated invoice_list to the render function
    doc.render({"name": name,
                "phone": phone,
                "invoice_list": invoice_list,
                "subtotal": subtotal,
                "salestax": str(salestax * 100) + "%",
                "total": total})

    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", "Invoice Complete")

    new_invoice()

# Create the main window
window = tkinter.Tk()
window.title("Invoice Generator Form")


#Create the GUI
frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

first_name_label = tkinter.Label(frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=0, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1, column=2)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

desc_label = tkinter.Label(frame, text="Description")
desc_label.grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

price_label = tkinter.Label(frame, text="Unit Price")
price_label.grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=3, column=2)

add_item_button = tkinter.Button(frame, text = "Add item", command = add_item)
add_item_button.grid(row=4, column=2, pady=5)

edit_save_button = tkinter.Button(frame, text="Edit Item", command=toggle_edit_save)
edit_save_button.grid(row=5, column=2, pady=5)

delete_item_button = tkinter.Button(frame, text="Delete Item", command=delete_item)
delete_item_button.grid(row=6, column=2, pady=5)

columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")

tree.grid(row=7, column=0, columnspan=3, padx=20, pady=10)

save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=8, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=9, column=0, columnspan=3, sticky="news", padx=20, pady=5)

# Start the main event loop
window.mainloop()