

selected_label_size = "100mm x 150mm"
selected_labels_per_page = 1

def apply_label_size():
    global selected_label_size
    global current_label_size_var

    selected_label_size = selected_label_size_var.get()

def apply_labels_per_page():

    global selected_labels_per_page
    selected_labels_per_page = int(selected_labels_per_page_var.get())
