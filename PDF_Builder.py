# --------------------------
# name: PDF Builder
# version: 1.0
# developer: Filippo Gatti
# --------------------------
import io
import os
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import pymupdf
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD

# CONSTANTS
valid_extensions = ['pdf', 'png', 'jpg', 'jpeg', 'gif', 'ico', 'tiff', 'bmp', 'pdn', 'heic']
IMG_WIDTH = 670
IMG_HEIGHT = 950
DIVIDER = 25

# VARIABLES
elements_of_listbox = []
images_for_canva = []
images_to_save = []


# PROCESSING FUNCTIONS
# create preview image
def create_image_preview(files_path):
    global images_for_canva, images_to_save

    for ii in range(len(images_for_canva), len(files_path)):
        el = files_path[ii]
        if type(el) is tuple:
            with pymupdf.open(el[0]) as doc:
                page = doc.load_page(el[1])
                pix = page.get_pixmap()
                img_b = pix.tobytes('ppm')
                with Image.open(io.BytesIO(img_b)) as img:
                    images_for_canva.append(ImageTk.PhotoImage(img.resize(
                        (IMG_WIDTH, IMG_HEIGHT), Image.Resampling.LANCZOS))
                    )
                    images_to_save.append(img)
        else:
            with Image.open(el) as img:
                images_for_canva.append(ImageTk.PhotoImage(img.resize(
                    (IMG_WIDTH, IMG_HEIGHT), Image.Resampling.LANCZOS))
                )
                images_to_save.append(img)


def convert_to_pdf():
    global images_to_save

    if images_to_save:  # if is not empty
        # get from user the name and the destination of the file
        name_file = filedialog.asksaveasfile(initialfile='merged',
                                             defaultextension='pdf',
                                             initialdir=r'C:\Users\{0}\Desktop'.format(os.getlogin()))

        if name_file is not None:
            images_to_save[0].save(name_file.name, save_all=True, append_images=images_to_save[1:])

            messagebox.showinfo(title='Informazioni', message='Salvataggio avvenuto con successo!')
            reset()

    else:
        messagebox.showerror(title='Errore', message='Nessuna immagine da salvare')


def reset():
    global elements_of_listbox, images_for_canva, images_to_save

    canva_view.delete('all')
    listbox.delete(0, tk.END)
    canva_scrollbar.pack_forget()
    mask_label.pack(fill=tk.BOTH, expand=True)  # put again the mask

    # disable sorting buttons
    sort_top_btn.state(['disabled'])
    sort_up_btn.state(['disabled'])
    sort_down_btn.state(['disabled'])
    sort_bottom_btn.state(['disabled'])
    
    # disable the remove and delete buttons
    remove_btn.state(['disabled'])
    delete_all_btn.state(['disabled'])
    convert_btn.state(['disabled'])

    elements_of_listbox = []
    images_for_canva = []
    images_to_save = []


# GUI FUNCTIONS
def scroll_to_selected_item(event):
    if listbox.curselection():
        i = listbox.curselection()[0]  # is a tuple
        offset = float((i * IMG_HEIGHT + i * DIVIDER) /
                       (len(images_for_canva) * IMG_HEIGHT + (len(images_for_canva) - 1) * DIVIDER))
        # y position of the clicked image divided by all images and all divider (one less cause last image hasn't it)
        canva_view.yview_moveto(offset)  # percentage value


def get_data_dropped(event):
    # print(event.data)
    docs = root.tk.splitlist(event.data)
    # row = event.data[1: -1]
    # docs = row.split("} {")

    add_elements(docs)


def update_canva():
    canva_view.delete('all')

    update_progressbar()
    populate_canva()


def populate_canva():
    global images_for_canva

    for img in images_for_canva:
        i = images_for_canva.index(img)
        canva_view.create_image(canva_view.winfo_width() / 2, i * IMG_HEIGHT + i * DIVIDER, anchor='n', image=img)
        
        # put a line as separator, but only if there are almost two elements and not at the top
        if len(images_for_canva) > 1 and i != 0:
            line_y_coord = i * IMG_HEIGHT + i * DIVIDER - DIVIDER / 2
            canva_view.create_line(0, line_y_coord, canva_view.winfo_width(), line_y_coord)

    canva_scrollbar_init()
    scroll_to_selected_item(event=None)
    update_progressbar()


def sorting(f):
    global elements_of_listbox, images_for_canva, images_to_save

    i = listbox.curselection()[0]  # is a tuple
    val = listbox.get(i)
    listbox.delete(i)

    if f == 'top':
        listbox.insert(0, val)
        elements_of_listbox.insert(0, elements_of_listbox.pop(i))  # update list with full path
        images_for_canva.insert(0, images_for_canva.pop(i))  # update list with images
        images_to_save.insert(0, images_to_save.pop(i))  # update list with images
        listbox.select_set(0)  # select item already selected
    elif f == 'bottom':
        listbox.insert(tk.END, val)
        elements_of_listbox.insert(len(elements_of_listbox), elements_of_listbox.pop(i))  # update list with full path
        images_for_canva.insert(len(images_for_canva), images_for_canva.pop(i))  # update list with images
        images_to_save.insert(len(images_to_save), images_to_save.pop(i))  # update list with images
        listbox.select_set(tk.END)  # select item already selected
    else:
        listbox.insert(i + f, val)
        elements_of_listbox.insert(i + f, elements_of_listbox.pop(i))  # update list with full path
        images_for_canva.insert(i + f, images_for_canva.pop(i))  # update list with images
        images_to_save.insert(i + f, images_to_save.pop(i))  # update list with images
        listbox.select_set(i + f)  # select item already selected

    update_canva()


def add_elements(new_doc):
    global elements_of_listbox

    if not new_doc:  # if it's empty
        new_doc = filedialog.askopenfilenames()

    # progressbar init
    progressbar.pack(side=tk.BOTTOM, anchor=tk.W, pady=5, padx=5)
    update_progressbar()

    for doc in new_doc:
        doc_ext = doc.split('.')[-1].lower()
        if doc_ext in valid_extensions:

            if doc_ext == 'pdf':
                pages = pymupdf.open(doc)

                for page in pages:
                    n = page.number
                    listbox.insert(tk.END, doc.split('/')[-1] + '-' + str(n+1))
                    elements_of_listbox.append((doc, n))

            else:
                listbox.insert(tk.END, doc.split('/')[-1])
                elements_of_listbox.append(doc)

        else:
            name = doc.split('/')[-1]
            messagebox.showerror(title="Tipo di file non valido",
                                 message=f"Il formato del documento {name} non è valido.")

    if mask_label.winfo_exists() and new_doc:
        mask_label.pack_forget()  # remove the label mask

        # disable sorting buttons
        sort_top_btn.state(['!disabled'])
        sort_up_btn.state(['!disabled'])
        sort_down_btn.state(['!disabled'])
        sort_bottom_btn.state(['!disabled'])

        # reactivate remove and delete and save buttons
        remove_btn.state(['!disabled'])
        delete_all_btn.state(['!disabled'])
        convert_btn.state(['!disabled'])

    check_for_scrollbar()
    update_progressbar()
    create_image_preview(elements_of_listbox)
    update_progressbar()
    update_canva()


def remove_elements():
    global elements_of_listbox, images_for_canva, images_to_save

    i = listbox.curselection()[0]  # is a tuple
    listbox.delete(i)
    elements_of_listbox.remove(elements_of_listbox[i])
    images_for_canva.remove(images_for_canva[i])
    images_to_save.remove(images_to_save[i])

    check_for_scrollbar()
    update_canva()

    if len(elements_of_listbox) == 0:
        mask_label.pack(fill=tk.BOTH, expand=True)  # put again the mask
        canva_scrollbar.pack_forget()

        # disable sorting buttons
        sort_top_btn.state(['disabled'])
        sort_up_btn.state(['disabled'])
        sort_down_btn.state(['disabled'])
        sort_bottom_btn.state(['disabled'])

        # disable the remove and delete and save buttons
        remove_btn.state(['disabled'])
        delete_all_btn.state(['disabled'])
        convert_btn.state(['disabled'])


def check_for_scrollbar():
    global elements_of_listbox

    if len(elements_of_listbox) > 33:
        lbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    else:
        lbox_scrollbar.pack_forget()


def update_progressbar():
    progressbar['value'] += 1
    print(progressbar['value'])
    progressbar.update_idletasks()

    if progressbar['value'] == 5:
        progressbar['value'] = 0
        progressbar.pack_forget()


def canva_scrollbar_init():
    canva_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canva_view.config(yscrollcommand=canva_scrollbar.set, scrollregion=canva_view.bbox('all'))
    canva_scrollbar.config(command=canva_view.yview)

    def _on_mousewheel(event):
        canva_view.yview_scroll(int(-1*(event.delta/120)), 'units')

    canva_view.bind_all('<MouseWheel>', _on_mousewheel)


# ROOT and settings
root = TkinterDnD.Tk()
root.title("PDF Builder")
root.iconbitmap('ico-64.ico')  # http://icons8.com/icons
width = root.winfo_screenwidth()
height = root.winfo_screenheight()
root.geometry(f'{int(width*0.8)}x{int(height*0.7)}+{int(width*0.1)}+{int(height*0.1)}')
root.resizable(False, False)

# STYLE
s = ttk.Style()
s.theme_use('xpnative')
s.configure('TButton', font=('Helvetica', 22, 'bold'), relief=tk.RAISED)
s.configure('TLabel', font=('Helvetica', 12, 'bold', 'italic'))
s.configure('mask.TLabel',
            font=('Helvetica', 20, 'bold'), background='white', foreground='grey',
            anchor=tk.CENTER, justify=tk.CENTER)
s.configure('progress.TLabel', font=('Helvetica', 10))

# basement frames
leftFrame = tk.Frame(root)
rightFrame = tk.Frame(root)
leftFrame.place(relheight=1, relwidth=0.3)
rightFrame.place(relx=0.3, relheight=1, relwidth=0.7)

# LABELS as title
ttk.Label(leftFrame, text="Ordina i documenti:").pack()
ttk.Label(rightFrame, text="Anteprima").pack()

# structure of the left frames
leftFrame1 = tk.Frame(leftFrame)
leftFrame2 = tk.Frame(leftFrame)
leftFrame1.place(rely=0.05, relheight=0.9, relwidth=0.85)
leftFrame2.place(rely=0.08, relx=0.85, relheight=0.9, relwidth=0.15)

# LISTBOX
listboxVar = tk.Variable()
listboxVar.set(elements_of_listbox)
listbox = tk.Listbox(leftFrame1, listvariable=listboxVar)
listbox.pack(padx=2, fill=tk.BOTH, expand=True)
# drop into listbox event
listbox.drop_target_register(DND_FILES)
listbox.dnd_bind('<<Drop>>', lambda e: get_data_dropped(e))
# select element of listbox event
listbox.bind('<<ListboxSelect>>', scroll_to_selected_item)

# MASK LABEL
mask_label = ttk.Label(listbox, text="Clicca qui\no\ntrascina qui\ni documenti", style='mask.TLabel')
mask_label.pack(fill=tk.BOTH, expand=True)
# mask label events
mask_label.bind('<Button>', lambda a: add_elements([]))
mask_label.drop_target_register(DND_FILES)
mask_label.dnd_bind('<<Drop>>', lambda e: get_data_dropped(e))

# ACTION BUTTONS
sort_top_btn = ttk.Button(leftFrame2, text="\u2B71", command=lambda: sorting('top'), state=tk.DISABLED)
sort_top_btn.pack(pady=5)
sort_up_btn = ttk.Button(leftFrame2, text="\u2B9D", command=lambda: sorting(-1), state=tk.DISABLED)
sort_up_btn.pack(pady=5)
sort_down_btn = ttk.Button(leftFrame2, text="\u2B9F", command=lambda: sorting(+1), state=tk.DISABLED)
sort_down_btn.pack(pady=5)
sort_bottom_btn = ttk.Button(leftFrame2, text="\u2B73", command=lambda: sorting('bottom'), state=tk.DISABLED)
sort_bottom_btn.pack(pady=5)
ttk.Button(leftFrame2, text="\u002B", command=lambda: add_elements([])).pack(pady=5)
remove_btn = ttk.Button(leftFrame2, text="\u2212", command=remove_elements, state=tk.DISABLED)
remove_btn.pack(pady=5)
delete_all_btn = ttk.Button(leftFrame2, text="x", command=lambda: reset(), state=tk.DISABLED)
delete_all_btn.pack(pady=5)
ttk.Button(leftFrame2, text="\u003F", command=lambda: os.startfile('guida.pdf')).pack(pady=5)
convert_btn = ttk.Button(leftFrame2, text="\u2713", command=lambda: convert_to_pdf(), state=tk.DISABLED)
convert_btn.pack(pady=30)

# PREVIEW CANVA
canva_view = tk.Canvas(rightFrame, background='white', relief=tk.RIDGE, border=2)
canva_view.pack(fill=tk.BOTH, expand=True)

# SCROLLBARS
lbox_scrollbar = tk.Scrollbar(listbox)
listbox.config(yscrollcommand=lbox_scrollbar.set)
lbox_scrollbar.config(command=listbox.yview)

canva_scrollbar = tk.Scrollbar(canva_view)

# PROGRESSBAR
progressbar = ttk.Progressbar(leftFrame, orient='horizontal', mode='determinate',
                              length=0.25 * root.winfo_width(), maximum=5, value=0)

# INIT
root.mainloop()
