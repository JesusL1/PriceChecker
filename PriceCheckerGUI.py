import tkinter as tk
from tkinter import ttk
from tkinter import font
import re
import PriceChecker as PC


HEIGHT = 350
WIDTH = 850
root = tk.Tk()
root.title('PRICE CHECKER APPLICATION - J')

def AddProduct(webLink, productPrice):
    """ Adds a product to the excel sheet once the arguments have been validated.
    Parses the website based on the arguments passed.
    """
    websiteLink_entry.delete(0,'end')  # clears the website entry field
    price_entry.delete(0, 'end')  # clears the price entry field
    productPrice = float(productPrice)
    PC.ValidateWebsite(webLink, productPrice, 'Add')

def InsertText(result, color):
    """ Inserts text to the tkinter GUI and color-codes it based on color variable
    Args: 
        result: text that will be inserted
        color: the background color of the text
    """
    print(result)
    output_text.insert(tk.END, '\n')
    if color == 0:
        output_text.insert(tk.END, result)
    elif color == 1:
        output_text.tag_config('worsePrice', background='#ffcce6')
        output_text.insert(tk.END, result, 'worsePrice')
    elif color == 2:
        output_text.tag_config('equalPrice', background='yellow')
        output_text.insert(tk.END, result, 'equalPrice')
    elif color == 3:
        output_text.tag_config('betterPrice', background='#4dff4d')
        output_text.insert(tk.END, result, 'betterPrice')
    elif color == 4:
        output_text.tag_config('newItem', background='#00cc66', underline=1)
        output_text.insert(tk.END, result, 'newItem')
    output_text.see('end')  # makes sure to move scrollbar to the last text inserted

def CheckIfWebsite(websiteLink):
    """ Checks if the input in the website link entry field is a valid website link.
    Args:
      websiteLink: the string value entered in the website link entry field
    Returns: 
        True if website link appears to be a valid website (matches the criteria of a website structure)
    """
    regex = re.compile(
        r'^(?:http|ftp)s?://' # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|' #domain...
        r'localhost|' # localhost...
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' # ...or ip
        r'(?::\d+)?' # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)
    if re.match(regex, websiteLink) is not None:
        return True
    else:
        return False

def CheckIfPrice(productPrice):
    """ Checks if the input in the price entry field is a valid product price.
    Args:
      productPrice: the string value entered in the price entry field
    Returns: 
        True if valid price, false otherwise
    """
    if productPrice == "" or float(productPrice) <= 0.02:  # user must enter a product price (can't be an empty string), price must be worth at least $0.02 to be considered a valid price
        return False
    else:
        return True

def Validator(*args):
    """ Validates the websiteLink and productPrice entry fields. The 'Add Button' will be clickable only when both entries meet
    the requirements. The entries are considered valid by their check functions.
    """
    websiteLink = CheckIfWebsite(websiteLink_entered.get())
    productPrice = CheckIfPrice(price_entered.get())

    if productPrice == False:
        valid_price_label.place_forget()
        invalid_price_label.place(relx=0.48,rely=0.35, relheight=0.08)
        add_button.config(state='disabled')
    if productPrice == True:
        invalid_price_label.place_forget()
        valid_price_label.place(relx=0.48,rely=0.35, relheight=0.08)
    if websiteLink == False:
        valid_websiteLink_label.place_forget()
        invalid_websiteLink_label.place(relx=0.76,rely=0.15, relheight=0.08)
        add_button.config(state='disabled')
    if websiteLink == True:
        invalid_websiteLink_label.place_forget()
        valid_websiteLink_label.place(relx=0.76,rely=0.15, relheight=0.08)
    if websiteLink == True and productPrice == True:
        add_button.config(state='normal')
   
def PriceValidate(action, index, value_if_allowed, prior_value, text, validation_type, trigger_type, widget_name):
    """ Restricts the entered input for a product's price. Only allows the user to enter numbers and a single decimal 
    otherwise a bell sound is placed.
    Note that this function has many parameters although only needing 'value_if_allowed'. I included all of them for
    future reference. Source: https://stackoverflow.com/questions/4140437/interactively-validating-entry-widget-content-in-tkinter/4140988#4140988
    """
    regex = re.compile(r'\d+(?:\.\d{0,2})?$')  # Checks for a decimal number with a max of 2 decimal places
    if value_if_allowed =="":  # allows for the entry field to be empty
        return True
    else:
        if value_if_allowed and re.match(regex, value_if_allowed) is not None:
            try:
                float(value_if_allowed)
                return True
            except ValueError:
                price_entry.bell()
                return False
        else:
            price_entry.bell()
            return False


""" Creates a canvas and a frame inside the canvas """
canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()
frame = tk.Frame(root, bg='#EDD4CE')
frame.place(relx=0, rely=0, relwidth=1, relheight=0.55)

""" Creates the variables that will hold the website link and product price
Also traces the variables while being passed/validated to the Validator function """
websiteLink_entered = tk.StringVar(root)
price_entered = tk.StringVar(root)
websiteLink_entered.trace("w", Validator)
price_entered.trace("w", Validator)

""" Creates labels for user to know where to enter the proper information """
websiteLink_label = tk.Label(frame,text='Enter Product Website Link:', bg='#EDD4CE', font='Helvetica 16 bold')
websiteLink_label.place(relx=0.01, rely=0.05, relheight=0.08)
price_label = tk.Label(frame,text='Enter Alert Price: $', bg='#EDD4CE', font='Helvetica 16 bold')
price_label.place(relx=0.01,rely=0.35, relheight=0.08)

""" Creates entries for user input. 
The price entry is validated interactively which restricts the user from entering an invalid key stroke. 
This requires the use of validatecommand which registers a validation function """
websiteLink_entry = tk.Entry(frame, bg='white', font=12, textvariable=websiteLink_entered)
websiteLink_entry.place(relx=0.01,rely=0.15, relwidth=0.75, relheight=0.08)
price_validate_helper = (frame.register(PriceValidate), '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
price_entry = tk.Entry(frame, validate="key", validatecommand=price_validate_helper, textvariable=price_entered, font='Helvetica 18')
price_entry.place(relx=0.23, rely=0.35, relwidth=0.25, relheight=0.08)

""" Creates labels for valid/invalid entries. Invalid labels are placed by default. """
valid_websiteLink_label = tk.Label(frame,text='VALID WEBSITE', bg='#66ff66', font='Helvetica 10 bold')
valid_price_label = tk.Label(frame,text='VALID PRICE', bg='#66ff66', font='Helvetica 10 bold')
invalid_websiteLink_label = tk.Label(frame,text='INVALID WEBSITE', bg='#ff3333', font='Helvetica 10 bold')
invalid_websiteLink_label.place(relx=0.76,rely=0.15, relheight=0.08)
invalid_price_label = tk.Label(frame,text='INVALID PRICE', bg='#ff3333', font='Helvetica 10 bold')
invalid_price_label.place(relx=0.48,rely=0.35, relheight=0.08)

""" Creates a button that is disabled by default. When it's enabled and clicked, it will call AddProduct() with the appropriate arguments """
add_button = tk.Button( frame, text='ADD A PRODUCT', bg='#4CAF50', borderwidth='3', font='Helvetica 12 bold', 
    cursor='hand2', state='disabled', command= lambda: AddProduct(websiteLink_entered.get(), price_entered.get()) )
add_button.place(relx=0.30, rely=0.65, relwidth=0.165, relheight=0.17)

""" Creates a button that when clicked will call CheckPrices() imported from PriceChecker.py """
check_button = tk.Button(frame, text='CHECK PRICES', bg='#ffff99', borderwidth='3', font='Helvetica 13 bold', cursor='hand2', command=PC.CheckPrices)
check_button.place(relx=0.50, rely=0.65, relwidth=0.165, relheight=0.17)

""" Creates a scrollbar and a text output box that is placed on the bottom of the window """
scrollb = tk.Scrollbar(root)
scrollb.pack(side=tk.RIGHT, fill=tk.Y)
output_text = tk.Text(root, height=20, width=110)
output_text.pack(side=tk.BOTTOM, fill=tk.Y)
scrollb.config(command=output_text.yview)
output_text.config(yscrollcommand=scrollb.set)
