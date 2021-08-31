import tkinter as tk
from tkinter import filedialog
from tkinter import *
from PIL import Image, ImageTk

import certificate as cer
import location
# ================================================== WINDOW ============================================================
class Window:
    def __init__(self):
    # INITIAL STRING VALUES == NULL
        self.spreadsheet_url = ''
        self.spreadsheet_path = ''

        self.template_path = ''
        self.font_style = ''
        self.font_colour = ''
        self.fontsize = ''
        self.coord = ''

        self.cutoff = ''

        self.emailid = ''
        self.apppass = ''
        self.email_subject = ''
        self.email_body = ''
    # INITIAL SPREADSHEET TYPE: GOOGLE
        self.spreadsheet_type = 1
    # INITIAL CUTOFF : NO
        self.cutoff_type = 2
    # INITIAL SEND EMAIL : YES
        self.email_type = 1

# ======= ROOT
        root = tk.Tk()

# ============================================== WINDOW SIZE (MAX & MIN) ==============================================
        root.geometry("800x604+0+0")
        # root.minsize(800, 615)
        # root.maxsize(1000, 1000)

# =================================== Esc & ENTER BUTTON BIND
        root.bind("<Escape>", lambda x: root.destroy())
        # root.bind('<Return>', lambda event=None: sub_btn.invoke())

# ================================================== TITLE =============================================================
        root.title('Auto Certificate Generator')

    # ======= SPREADSHEET DETAILS VARIABLE
        Spreadsheet_Type_var = tk.IntVar()
        spreadsheet_url_var = tk.StringVar()
        spreadsheet_path_var = tk.StringVar()

    # ======= TEMPLATE DETAILS VARIABLE
        template_path_var = tk.StringVar()
        font_type_var = tk.StringVar()
        font_colour_var = tk.StringVar()
        font_size_var = tk.StringVar()
        coord_var = tk.StringVar()

    # ======= CUTOFF & EMAIL (YES/NO) VARIABLE
        cutoff_Type_var = tk.IntVar()
        cutoff_var = tk.StringVar()
        Email_YesNo_var = tk.IntVar()

    # ======= EMAIL DETAILS VARIABLE
        emailid_var = tk.StringVar()
        apppass_var = tk.StringVar()

# ======= SUBMIT FUNCTION (WITH ERROR BOX)
        def submit():
            self.spreadsheet_url = str(spreadsheet_url_var.get())
            self.spreadsheet_path = str(spreadsheet_path_var.get())

            self.template_path = str(template_path_var.get())
            self.font_style = str(font_type_var.get())
            self.font_colour = str(font_colour_var.get())
            self.fontsize = str(font_size_var.get())
            self.coord = str(coord_var.get())

            self.cutoff = str(cutoff_var.get())

            self.emailid = str(emailid_var.get())
            self.apppass = str(apppass_var.get())
            self.email_subject = str(email_subject_txt.get("1.0", "end-1c"))
            self.email_body = str(email_body_txt.get("1.0", "end-1c"))

            self.spreadsheet_type = int(Spreadsheet_Type_var.get())
            self.cutoff_type = int(cutoff_Type_var.get())
            self.email_type = int(Email_YesNo_var.get())

    # ====== COUNT NUMBERS OF BLANK STRINGS
            count = [self.template_path, self.font_style, self.font_colour, self.fontsize, self.coord].count('')
            countmail = [self.emailid, self.apppass].count('')

# ============ ERROR BOX (TOP LEVEL)
            if count >= 1 or (self.spreadsheet_type == 1 and self.spreadsheet_url == '') or (
                    self.spreadsheet_type == 2 and self.spreadsheet_path == '') or (
                    self.cutoff_type == 1 and self.cutoff == '') or (self.email_type == 1 and countmail >= 1):
                error_window = Toplevel(root)
    # ====== ERROR BOX LOCATION
                error_window.geometry("200x100+203+45")
                error_window.minsize(200, 100)
                error_window.maxsize(200, 100)
    # ====== ERROR BOX LABEL
                Label(error_window, text="⚠️ *Required field is empty", font=('Times', 12, 'bold'), fg="Red", padx=30,
                      pady=90).pack()
# ====== DESTROY WINDOW
            else:
                root.destroy()

# ===== BROWSE BUTTON : SPREADSHEET PATH
        def spreadsheet_browsefunc():
            filename = filedialog.askopenfilename(
                filetypes=(("excel files", "*.xlsx"), ("csv files", "*.csv"), ("All files", "*.*")))
            spreadsheet_path_entry.insert(END, filename)

# ===== BROWSE BUTTON : TEMPLATE PATH
        def template_browsefunc():
            filename = filedialog.askopenfilename(
                filetypes=(("jpeg files", "*.jpeg"), ("png files", "*.png"), ("All files", "*.*")))
            template_path_entry.insert(END, filename)

# ===== LOCATE BUTTON
        def locate_button():
            self.template_path = str(template_path_var.get())
    # === COLLECT COORDINATES USING --> FUNCTION: MAIN & FILE: LOCATION.PY
            X,Y=location.main(self.template_path)
    # === COORD STRING (X,Y)
            self.coord=str(X)+","+str(Y)
    # === DELETE PREVIOUS VALUES FROM COORD ENTRY BOX
            coord_entry.delete(0,END)
    # === UPDATE NEW COORD STRING IN COORD ENTRY BOX
            coord_entry.insert(0,self.coord)

# ===== PREVIEW BUTTON
        def preview_button():
            self.template_path = str(template_path_var.get())
            self.font_style = str(font_type_var.get())
            self.font_colour = str(font_colour_var.get())
            self.fontsize = str(font_size_var.get())
            self.coord = str(coord_var.get())
    # === CREATE A PREVIEW CERTIFICATE FROM GIVEN TEMPLATE VALUES --> FUNCTION: CREATECERTIFICATE & FILE: PREVIEW.PY
            cer.previewcertificate(self.template_path,self.font_style,self.font_colour,self.fontsize,self.coord)

    # === OPEN SAMPLE CERTIFICATE WITH GIVEN TEMPLATE VALUES
            global image,img
            image=Image.open("sample_image.jpeg")
    # === ORIGINAL WIDTH & HEIGHT OF PREVIEW CERTIFICATE
            width, height = image.size
    # === 1/3 (WIDTH & HEIGHT) OF PREVIEW CERTIFICATE
            new_width = int(width * 0.33)
            new_height = int(height * 0.33)
    # === RESIZE PREVIEW CERTIFICATE WITH 1/3 (WIDTH & HEIGHT)
            image = image.resize((new_width, new_height), Image.ANTIALIAS)
            string_geo = str(new_width) + "x" + str(new_height) + "+800+0"
            img = ImageTk.PhotoImage(image)

    # === CREATE A TOPLEVEL WIDGET
            top = Toplevel(root)
    # === OPEN TOPLEVEL WIDGET SIZE--> SIZE: W = 1/3 (WIDTH OF ORIGINAL IMAGE) x H = 1/3 (HEIGHT OF ORIGINAL IMAGE)
        # ================ POSITION--> FROM : W=800 + H=0
            top.geometry(string_geo)
    # === PACK TOPLEVEL WINDOW WITH 1/3 PREVIEW CERTIFICATE
            panel = Label(top, image=img)
            panel.pack(side="bottom")

# ===== SPREADSHEETS TYPE : LOCAL
        def disableURL():
            spreadsheet_url_entry.configure(state="disabled")
            spreadsheet_path_entry.configure(state="normal")
            spreadsheet_path_button.configure(state="normal")
            spreadsheet_url_entry.update()
            spreadsheet_path_entry.update()
            spreadsheet_path_button.update()

# ===== SPREADSHEETS TYPE : GOOGLE
        def disablePATH():
            spreadsheet_path_entry.configure(state="disabled")
            spreadsheet_path_button.configure(state="disabled")
            spreadsheet_url_entry.configure(state="normal")
            spreadsheet_path_entry.update()
            spreadsheet_url_entry.update()
            spreadsheet_path_button.update()

# ===== CUTOFF : YES
        def enableCutoff():
            cutoff.configure(state="normal")
            cutoff.update()

# ===== CUTOFF : NO
        def disableCutoff():
            cutoff.configure(state="disabled")
            cutoff.update()

# ===== EMAIL : YES
        def enableEmail():
            emailid_entry.configure(state="normal")
            app_password_entry.configure(state="normal")
            email_subject_txt.configure(state="normal")
            email_body_txt.configure(state="normal")
            emailid_entry.update()
            app_password_entry.update()
            email_subject_txt.update()
            email_body_txt.update()

# ===== EMAIL : NO
        def disableEmail():
            emailid_entry.configure(state="disabled")
            app_password_entry.configure(state="disabled")
            email_subject_txt.configure(state="disabled")
            email_body_txt.configure(state="disabled")
            emailid_entry.update()
            app_password_entry.update()
            email_subject_txt.update()
            email_body_txt.update()

# ====== BACKGROUND COLOUR : ALL FRAMES
        bg_color = "black"
# =============================================== TITLE FRAME ==========================================================
        title = Label(root, text="Auto Certificate Generator", bd=10, relief=GROOVE, bg=bg_color, fg="white",
                      font=("Arial", 18, "bold"), pady=2)
        title.pack(fill=X)

# ================================================= FRAME 1 ============================================================
        F1 = LabelFrame(root, bd=10, relief=GROOVE, text="Spreadsheet Details", bg=bg_color, fg="gold",
                        font=("Arial", 13, "bold"))
        F1.place(x=0, y=50, relwidth=1)

# ========== SPREADSHEET TYPE LABEL
        spreadsheet_type_lbl = Label(F1, bg=bg_color, fg="white", text="Spreadsheet Type", font=("Arial", 13, "bold"))
        spreadsheet_type_lbl.grid(row=0, column=0, padx=10, pady=10)

# ========== SPREADSHEET TYPE RADIOBUTTON : (GOOGLE / LOCAL)
        spreadsheet_type1_Radiobutton = Radiobutton(F1, text='Google Sheets', variable=Spreadsheet_Type_var, value=1,
                                                    command=disablePATH)
        spreadsheet_type1_Radiobutton.grid(row=0, column=1, padx=13, pady=10, sticky="nsew")
        spreadsheet_type2_Radiobutton = Radiobutton(F1, text='Local', variable=Spreadsheet_Type_var, value=2,
                                                    command=disableURL)
        spreadsheet_type2_Radiobutton.grid(row=0, column=2, padx=13, pady=10, sticky="nsew")
# DEFAULT : GOOGLE
        Spreadsheet_Type_var.set(1)

# ========== SPREADSHEET URL LABEL
        spreadsheet_url_lbl = Label(F1, bg=bg_color, fg="white", text="Spreadsheet URL", font=("Arial", 13, "bold"))
        spreadsheet_url_lbl.grid(row=1, column=0, padx=10, pady=10)

# ========== SPREADSHEET URL ENTRY BOX
        spreadsheet_url_entry = Entry(F1, textvariable=spreadsheet_url_var, width=18, font="arial 13", bd=3,
                                      relief=SUNKEN)
        spreadsheet_url_entry.grid(row=1, column=1, padx=5, pady=5)

# ========== SPREADSHEET PATH LABEL
        spreadsheet_path_lbl = Label(F1, bg=bg_color, fg="white", text="Spreadsheet Path", font=("Arial", 13, "bold"))
        spreadsheet_path_lbl.grid(row=1, column=2, padx=5, pady=5)

# ========== SPREADSHEET PATH ENTRY BOX
        spreadsheet_path_entry = Entry(F1, textvariable=spreadsheet_path_var, width=18, font="arial 13", bd=3,
                                       relief=SUNKEN)
        spreadsheet_path_entry.grid(row=1, column=3, padx=5, pady=5)
# DEFAULT : DISABLED
        spreadsheet_path_entry.config(state='disabled')

# ========== SPREADSHEET PATH BROWSE BUTTON
        spreadsheet_path_button = Button(F1, text="Browse", command=spreadsheet_browsefunc, bd=4,
                                         font=("Arial", 10, "bold"))
        spreadsheet_path_button.grid(row=1, column=4, padx=10, pady=10, sticky="nsew")
# DEFAULT : DISABLED
        spreadsheet_path_button.configure(state="disabled")

# ================================================= FRAME 2 ============================================================
        F2 = LabelFrame(root, bd=10, relief=GROOVE, text="Template Details", bg=bg_color, fg="gold",
                        font=("Arial", 13, "bold"))
        F2.place(x=0, y=160, relwidth=1)

# ========== TEMPLATE PATH LABEL
        template_path_lbl = Label(F2, bg=bg_color, fg="white", text="Template Path", font=("Arial", 13, "bold"))
        template_path_lbl.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# ========== TEMPLATE PATH ENTRY BOX
        template_path_entry = Entry(F2, textvariable=template_path_var, width=18, font="arial 13", bd=3, relief=SUNKEN)
        template_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

# ========== TEMPLATE PATH BROWSE BUTTON
        template_path_button = Button(F2, text="Browse", command=template_browsefunc, bd=4, font=("Arial", 10, "bold"))
        template_path_button.grid(row=0, column=2, padx=5, pady=5)

# ========== FONT STYLE LABEL
        font_type_lbl = Label(F2, bg=bg_color, fg="white", text="Font Style", font=("Arial", 13, "bold"))
        font_type_lbl.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# ========== FONT STYLE OPTION MENU
        font_options = ["Arial Unicode.ttf", "Calibri.ttf", "Font 3", "Font 4", "Font 5"]

        font_type_optionmenu = OptionMenu(F2, font_type_var, *font_options)
        font_type_optionmenu.grid(row=1, column=1, padx=13, pady=10, sticky="nsew")
# DEFAULT : ARIAL UNICODE
        font_type_var.set("Arial Unicode.ttf")

# ========== FONT COLOUR LABEL
        font_colour_lbl = Label(F2, bg=bg_color, fg="white", text="Font Colour", font=("Arial", 13, "bold"))
        font_colour_lbl.grid(row=1, column=2, padx=10, pady=10, sticky="nsew")

# ========== FONT COLOUR OPTION MENU
        font_colour_options = ["Black", "White", "Red", "Lime", "Blue", "Yellow"]

        font_colour_type_optionmenu = OptionMenu(F2, font_colour_var, *font_colour_options)
        font_colour_type_optionmenu.grid(row=1, column=3, padx=13, pady=10, sticky="nsew")
# DEFAULT : BLACK
        font_colour_var.set("Black")

# ========== FONT SIZE LABEL
        font_type_lbl = Label(F2, bg=bg_color, fg="white", text="Font Size", font=("Arial", 13, "bold"))
        font_type_lbl.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

# ========== FONT SIZE SPINBOX
        font_size = Spinbox(F2, from_=0, to=200, increment=2, textvariable=font_size_var)
        font_size.grid(row=2, column=1, padx=13, pady=10, sticky="nsew")
# DEFAULT SIZE : 80
        font_size_var.set("80")

# ========== FONT LOCATION (COORDINATES) LABEL
        font_coordinates_lbl = Label(F2, bg=bg_color, fg="white", text="Font Coordinates", font=("Arial", 13, "bold"))
        font_coordinates_lbl.grid(row=2, column=2, padx=10, pady=10, sticky="nsew")

# ========== FONT LOCATION (COORDINATES) ENTRY BOX
        coord_entry = Entry(F2, textvariable=coord_var, width=18, font="arial 13", bd=3, relief=SUNKEN)
        coord_entry.grid(row=2, column=3, padx=5, pady=5, sticky="nsew")

# ========== FONT LOCATION (COORDINATES) PREVIEW BUTTON --> OPENCV
        coord_button = Button(F2, text='Locate', bd=4, font=("Arial", 10, "bold"),command=locate_button)
        coord_button.grid(row=2, column=4, padx=10, pady=10)

# ================================================= FRAME 3 ============================================================
        F3 = LabelFrame(root, bd=10, relief=GROOVE, text="Cutoff & Email", bg=bg_color, fg="gold",
                        font=("Arial", 13, "bold"))
        F3.place(x=0, y=320, relwidth=1)

# ========== CUTOFF LABEL
        cutoff_lbl = Label(F3, bg=bg_color, fg="white", text="Do you want to set Cutoff?", font=("Arial", 13, "bold"))
        cutoff_lbl.grid(row=0, column=0, padx=10, pady=10)

# ========== CUTOFF RADIOBUTTON (YES/NO)
        cutoff_type1_Radiobutton = Radiobutton(F3, text='Yes', variable=cutoff_Type_var, value=1, command=enableCutoff)
        cutoff_type1_Radiobutton.grid(row=0, column=1, padx=13, pady=10, sticky="nsew")
        cutoff_type2_Radiobutton = Radiobutton(F3, text='No', variable=cutoff_Type_var, value=2, command=disableCutoff)
        cutoff_type2_Radiobutton.grid(row=0, column=2, padx=13, pady=10, sticky="nsew")
# DEFAULT : NO
        cutoff_Type_var.set(2)

# ========== CUTOFF SPINBOX
        cutoff = Spinbox(F3, from_=0, to=100, increment=5, textvariable=cutoff_var)
        cutoff.grid(row=0, column=3, padx=13, pady=10)
# DEFAULT SIZE: 80
        cutoff_var.set("80")
# DEFAULT : DISABLE
        cutoff.configure(state="disabled")

# ========== EMAIL LABEL
        Email_YesNo_lbl = Label(F3, bg=bg_color, fg="white", text="Do you Want to send Emails?",
                                font=("Arial", 13, "bold"))
        Email_YesNo_lbl.grid(row=1, column=0, padx=10, pady=10)

# ========== MAIL (YES/NO) RADIOBUTTON
        Email_Yes_Radiobutton = Radiobutton(F3, text='Yes', variable=Email_YesNo_var, value=1, command=enableEmail)
        Email_Yes_Radiobutton.grid(row=1, column=1, padx=13, pady=10, sticky="nsew")
        Email_No_Radiobutton = Radiobutton(F3, text='No', variable=Email_YesNo_var, value=2, command=disableEmail)
        Email_No_Radiobutton.grid(row=1, column=2, padx=13, pady=10, sticky="nsew")
# DEFAULT : YES
        Email_YesNo_var.set(1)

# ================================================= FRAME 4 ============================================================

        F4 = LabelFrame(root, bd=10, relief=GROOVE, text="Email Details", bg=bg_color, fg="gold",
                        font=("Arial", 13, "bold"))
        F4.place(x=0, y=439, relwidth=1)

# ========== EMAIL_ID LABEL
        emailid_lbl = Label(F4, bg=bg_color, fg="white", text="Senders Email Id", font=("Arial", 13, "bold"))
        emailid_lbl.grid(row=0, column=0, padx=10, pady=10)
# ========== EMAIL_ID ENTRY BOX
        emailid_entry = Entry(F4, textvariable=emailid_var, width=18, font="Courier 13", bd=3, relief=SUNKEN,show='*')
        emailid_entry.grid(row=0, column=1, padx=5, pady=5)

# ========== EMAIL LOGIN APP PASSWORD LABEL
        app_password_lbl = Label(F4, bg=bg_color, fg="white", text="App Password", font=("Arial", 13, "bold"))
        app_password_lbl.grid(row=0, column=2, padx=10, pady=10)
# ========== EMAIL LOGIN APP PASSWORD ENTRY BOX
        app_password_entry = Entry(F4, textvariable=apppass_var, width=18, font="Courier 13", bd=3, relief=SUNKEN,show='*')
        app_password_entry.grid(row=0, column=3, padx=5, pady=5)

# ========== EMAIL SUBJECT LABEL
        email_subject_lbl = Label(F4, bg=bg_color, fg="white", text="Email Subject", font=("Arial", 13, "bold"))
        email_subject_lbl.grid(row=1, column=0, padx=10, pady=10)
# ========== EMAIL SUBJECT TEXT BOX
        email_subject_txt = Text(F4, width=20, height=3, font="arial 13", bd=3,
                                 relief=SUNKEN)
        email_subject_txt.grid(row=1, column=1, padx=5, pady=5)

# ========== EMAIL BODY LABEL
        email_body_lbl = Label(F4, bg=bg_color, fg="white", text="Email Text", font=("Arial", 13, "bold"))
        email_body_lbl.grid(row=1, column=2, padx=10, pady=10)
# ========== EMAIL BODY TEXT BOX
        email_body_txt = Text(F4, width=30, height=3, font="arial 13", bd=3, relief=SUNKEN)
        email_body_txt.grid(row=1, column=3, padx=5, pady=5)

# ========== PREVIEW BUTTON --> PIL
        preview_button = Button(F4, text='Preview Certificate', bd=4, font=("Arial", 10, "bold"),command=preview_button)
        preview_button.grid(row=2, column=3)

# ========== SUBMIT BUTTON
        submit_button = Button(F4, text='Submit', command=submit, bd=4, font=("Arial", 10, "bold"))
        submit_button.grid(row=2, column=4)

        root.mainloop()
