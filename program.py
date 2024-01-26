import win32com.client

import pyautogui
import win32gui 

import os
import time

import tkinter as tk
from tkinter import messagebox, filedialog
import uuid

# prior to using the app add Inverse, 
# Select Subject, and Fill With White to Default Actions

# data
background_image_path = r"C:\Users\skyma\Desktop\DSC09088.jpg"
overlay_image_path = r"C:\Users\skyma\Desktop\50zx.png"

def startProgram(options): 
    id = str(uuid.uuid4())
    hardcopies = options["hardcopies"]
    digital = options["digital"]
    
    # 1. access photoshop
    app = win32com.client.Dispatch("Photoshop.Application") 
    save_options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")

    # 2. open the main photo file
    background_doc = app.Open(options["image_path"])

    # 3. open the overlay png image in another tap and copy
    app.Open(options["overlay_path"])
    app.ActiveDocument.Selection.Copy()

    # 4. go back to previous tab and paste the overlay png image
    app.ActiveDocument = background_doc
    background_doc.Paste()     

    # 5. simulate Ctrl + T key-press to enable Free Transform tool
    def get_photoshop_window():
        # selects the photoshop window
        return win32gui.FindWindow("Photoshop", None)

    def bring_to_front(window):
    # brings photoshop window to front
        win32gui.ShowWindow(window, 9)
        win32gui.SetForegroundWindow(window)

    def save_as_jpeg(hardcopies):
        save_options.Quality = 12 # high quality

        if hardcopies:
             filename = "canada_passport_2_prints-" + id + ".jpg"
        else:
             filename = "canada_passport_digital-" + id + ".jpg"

        desktop_path = os.path.expanduser('~/Desktop')
        output_path = r"{}\{}".format(desktop_path, filename)  

        app.ActiveDocument.SaveAs(output_path, save_options, True)

    def zoom_in(num):
        for x in range(num):
            pyautogui.hotkey('ctrl', '=')
            pyautogui.hotkey('ctrl', '=') 
            pyautogui.hotkey('ctrl', '=') 

    def free_transform():
        pyautogui.hotkey('ctrl', 't') # simulate ctrl + t

    def center_image(doc): 
        doc.ArtLayers["Background"].Visible = False
        
        app.ActiveDocument.MergeVisibleLayers()

        doc.ArtLayers["Background"].Visible = True
        
        width_inner = abs(doc.ArtLayers["Background copy"].Bounds[0] - doc.ArtLayers["Background copy"].Bounds[2])
        height_inner = abs(doc.ArtLayers["Background copy"].Bounds[1] - doc.ArtLayers["Background copy"].Bounds[3])

        width_outer = abs(doc.ArtLayers["Background"].Bounds[0] - doc.ArtLayers["Background"].Bounds[2])
        height_outer = abs(doc.ArtLayers["Background"].Bounds[1] - doc.ArtLayers["Background"].Bounds[3])

        move_x = int(abs(width_outer - width_inner) / 2)
        move_y = int(abs(height_outer - height_inner) / 2)

        doc.ArtLayers["Background copy"].Translate(move_x, move_y)
        
        app.ActiveDocument.MergeVisibleLayers() # merge all the layers

    def duplicate_image(doc):
        return doc.ArtLayers["Background copy"].Duplicate()

    def translate_image_right(layer, pixels):
        app.Preferences.RulerUnits = 1 # set PIXELS as ruler units

        width = abs(layer.Bounds[0] - layer.Bounds[2])

        layer.Translate(width + pixels, 0)

    def create_image_stroke(doc):
        stroke_color = win32com.client.Dispatch("Photoshop.SolidColor") # create stroke color

        stroke_color.CMYK.Cyan = 0
        stroke_color.CMYK.Magenta = 0
        stroke_color.CMYK.Yellow = 0
        stroke_color.CMYK.Black = 100

        doc.Selection.Stroke(stroke_color, 1, 1, 2, 100, False) # 3 for outside stroke, 2 for color blending mode

    def create_doc(width_inches, height_inches, ppi, name):
        app.Preferences.RulerUnits = 2 # set INCHES as ruler units
        return app.Documents.Add(width_inches, height_inches, ppi, name)

    def copy_paste_image(from_doc, to_doc):
        app.ActiveDocument = from_doc
        from_doc.ArtLayers["Background"].Duplicate(to_doc)
        app.ActiveDocument = to_doc

    def resize_image(width_mm, height_mm):
        app.Preferences.RulerUnits = 4 # set MM as ruler units
        app.ActiveDocument.ResizeImage(width_mm, height_mm, 300, 8)

    def whiten_background(): 
        app.DoAction("Select Subject", "Default Actions")
        app.DoAction("Inverse", "Default Actions")
        app.DoAction("Fill With White", "Default Actions")

        app.ActiveDocument.Selection.Deselect()

    def increase_brightness(doc, percent):
        brightness_layer = doc.ArtLayers.Add()
        brightness_layer.name = "Brightness Adjustment"
        doc.ArtLayers["Background"].AdjustBrightnessContrast(percent, 1)

        app.ActiveDocument.MergeVisibleLayers()

    def crop_by_overlay(doc):
        overlay_layer = doc.ArtLayers["Layer 1"] # get layer 1 

        crop_x1, crop_y1, crop_x2, crop_y2 = overlay_layer.Bounds # get rectangular bounds of layer 1

        app.ActiveDocument.Crop([crop_x1, crop_y1, crop_x2, crop_y2], 0, 0, 0) # crop the image

        overlay_layer.Delete() # delete the png overlay image

    def confirm():
        pyautogui.hotkey('enter') # simulate enter key press

    def display_popup():
        root = tk.Tk()
    
        root.withdraw()

        result = messagebox.askokcancel("Image Resizing Instruction", "Resize the png image" + 
                                    " so it looks similar to the image below " + 
                                    " only resize it by adjusting the vertex highlighted by green color"
                                    + " when done press OK to let the program proceed")
    
        if result: # user clicks ok button
            # 6. crop the image and remove the overlay
            photoshop_window = get_photoshop_window() 
        
            bring_to_front(photoshop_window)

            confirm()

            crop_by_overlay(background_doc)

            # 7. whiten the background
            whiten_background()
        
            # 8. increase the brightness level
            increase_brightness(background_doc, options["brightness"])

            # 9. resize image 50 by 70 mm 300 ppi
            resize_image(50, 70)

            if (digital):
                save_as_jpeg(False)

            if (hardcopies == False): 
                messagebox.showinfo("Success", "Image(s) Were Successfully prepared and saved to Desktop")
                return

            # 10. create 6x4 document
            final_doc = create_doc(options["width_inches"], options["height_inches"], 300, "final")

            # 11. copy and paste image from the previous document to the final one
            copy_paste_image(background_doc, final_doc)

            # 12. add stroke to the image
            create_image_stroke(final_doc)

            # 13. duplicate image and translate it to the right
            copied_layer = duplicate_image(final_doc)
            translate_image_right(copied_layer, 50)

            # 14. center the image
            center_image(final_doc)

            # 15. save to Desktop
            if (hardcopies):
                save_as_jpeg(True)

            messagebox.showinfo("Success", "Image(s) Were Successfully prepared and saved to Desktop")
    
        # wait
    
    time.sleep(0.5)

    # get photoshop window
    photoshop_window = get_photoshop_window()

    if photoshop_window:
        bring_to_front(photoshop_window)

        free_transform()

        zoom_in(3)
    
        display_popup() # displays pop up window to let the user
                    # adjust the size of the picture
    else:
        messagebox.showerror("Error", "Photoshop Window Not Found")
        print("Photoshop Window Not Found.")

# GUI
def open_file(text_box):
    file = filedialog.askopenfile(mode='r', filetypes=[('Image Files', '*.jpg; *.png; *.jpeg'),])
    
    if (file):
        filepath = os.path.abspath(file.name)
        text_box.insert(0, str(filepath))

def toggle_paper_size(label, width, height, obj):
    if (not obj["show_paper_size"]):
        label.grid(column=0, row=9, padx=20, pady=0, sticky='w')
        width.grid(column=0, row=10, padx=20, pady=5, sticky='w')
        height.grid(column=1, row=10, padx=20, pady=5, sticky='w')
    else:
        label.grid_remove()
        width.grid_remove()
        height.grid_remove()
    
    obj["show_paper_size"] = not obj["show_paper_size"]

def display_gui():
    main_window = tk.Tk()

    main_window.title("Skymart Photo Editor App (v 1.0)")

    label = tk.Label(main_window, text="Select Parameters and Click Create", foreground="#15678A", font=('Arial', 18, 'underline'))
    label.grid(column=0, row=0, padx=20, pady=20, sticky="w")

    label_image = tk.Label(main_window, text="1. Select an Image to Edit", foreground="black", font=('Arial', 12))
    label_image.grid(column=0, row=1, padx=20, pady=0, sticky='w')
 
    text_box_image = tk.Entry(main_window, text="Image Path", foreground="black", font=('Arial', 12), width=50)
    text_box_image.grid(column=0, row=2, padx=20, pady=5, sticky='w')

    btn_image_path = tk.Button(main_window, text="Browse Image", font=('Arial', 12), command=lambda: [ open_file(text_box_image) ])
    btn_image_path.grid(column=1, row=2, padx=20, pady=5, sticky='w')

    label_overlay = tk.Label(main_window, text="2. Select an Overlay Mask to Use", foreground="black", font=('Arial', 12))
    label_overlay.grid(column=0, row=3, padx=20, pady=0, sticky='w')

    text_box_overlay = tk.Entry(main_window, text="Overlay Path", foreground="black", font=('Arial', 12), width=50)
    text_box_overlay.grid(column=0, row=4, padx=20, pady=5, sticky='w')

    btn_overlay_path = tk.Button(main_window, text="Browse Overlay", font=('Arial', 12), command=lambda: [ open_file(text_box_overlay) ])
    btn_overlay_path.grid(column=1, row=4, padx=20, pady=5, sticky='w')

    label_brightess = tk.Label(main_window, text="3. Select Brightness Level (-100 to 100)", foreground="black", font=('Arial', 12))
    label_brightess.grid(column=0, row=5, padx=20, pady=0, sticky='w')

    text_box_brightess = tk.Entry(main_window, foreground="black", font=('Arial', 12), width=5)
    text_box_brightess.grid(column=0, row=6, padx=20, pady=5, sticky='w')
    text_box_brightess.insert(0, "20") # default value

    label_output_type = tk.Label(main_window, text="4. Select Output Type", foreground="black", font=('Arial', 12))
    label_output_type.grid(column=0, row=7, padx=20, pady=0, sticky='w')

    digital_bool = tk.BooleanVar()
    hardcopies_bool = tk.BooleanVar()
    obj = {"show_paper_size": False}
    
    checkbox_digital = tk.Checkbutton(main_window, text="Digital", font=('Arial', 12), variable=digital_bool, onvalue=True, offvalue=False)
    checkbox_digital.grid(column=0, row=8, padx=20, pady=0, sticky='w')

    label_paper_size = tk.Label(main_window, text="5. Enter the Paper Size, Width and Height in Inches", foreground="black", font=('Arial', 12))
    text_box_width = tk.Entry(main_window, foreground="black", font=('Arial', 12), width=3)
    text_box_heihgt = tk.Entry(main_window, foreground="black", font=('Arial', 12), width=3)

    text_box_width.insert(0, "6")
    text_box_heihgt.insert(0, "4")

    checkbox_prints = tk.Checkbutton(main_window, text="2 Prints", font=('Arial', 12), variable=hardcopies_bool, onvalue=True, offvalue=False, command=lambda: [toggle_paper_size(label_paper_size, text_box_width, text_box_heihgt, obj)])
    checkbox_prints.grid(column=1, row=8, padx=20, pady=0, sticky='w')

    def get_input(): 
        options = {
            "image_path": text_box_image.get(),
            "overlay_path": text_box_overlay.get(),
            "brightness": int(text_box_brightess.get()),
            "digital": digital_bool.get(),
            "hardcopies": hardcopies_bool.get(),
            "width_inches": int(text_box_width.get()),
            "height_inches": int(text_box_heihgt.get())
        }

        return options

    button_create = tk.Button(main_window, text="Create Photo(s)", font=('Arial', 14), background="#15678A", foreground="white", command=lambda: [startProgram(get_input())])
    button_create.grid(column=0, row=11, ipadx=10, ipady=10, padx=20, pady=20, sticky="w")

    main_window.mainloop()

def display_instructions():
    instructions_window = tk.Tk()
    
    instructions_window.geometry("300x300")
    instructions_window.title("Skymart Photo Editor App (v 1.0)")

    label = tk.Label(instructions_window, text="Istructions Prior to Use", foreground="#15678A", font=('Arial', 18, 'underline'))
    label.pack(padx=20, pady=20)

    instructions_label = tk.Label(instructions_window, text="1. Open Photoshop\n2. Sign In\n3. Close All the Tabs", foreground="black", font=('Arial', 14, 'italic'), justify="left")
    instructions_label.pack(padx=20, pady=0, side="top")

    continue_button = tk.Button(instructions_window, text="Continue", font=('Arial', 14), background="#15678A", foreground="white", command=lambda: [instructions_window.destroy(), display_gui()])
    continue_button.pack(padx=20, pady=20, side="right")

    instructions_window.mainloop()

display_instructions()






