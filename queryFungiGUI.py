import os
import re
import sys
import time
import pandas as pd

from tkinter import *
from tkinter import ttk, scrolledtext, messagebox
from tkinter.filedialog import asksaveasfilename

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

root = Tk()
root.title('QueryFungi')
root_path = os.getcwd().replace('\\', '\\\\')

root.iconphoto(True, PhotoImage(file=root_path + '\\mushroom.png'))
root.geometry('600x500+100+100')
root_theme = ttk.Style()
root_theme.theme_use('alt')

mainframe = ttk.Frame(root, padding='3 3 12 12')
mainframe.pack(fil=BOTH, expand=True)
mainframe.rowconfigure(8, weight=1)
mainframe.columnconfigure(0, weight=1)


# Table Windows
def table_window(table_list):
    root_table = Tk()
    root_table.title('Table')
    root_table.geometry('900x500+200+100')
    root_table_theme = ttk.Style()
    root_table_theme.theme_use('alt')

    table_frame = ttk.Frame(root_table, padding='3 3 12 12')
    table_frame.pack(fill=BOTH, expand=True)

    xscroll = Scrollbar(table_frame, orient=HORIZONTAL)
    yscroll = Scrollbar(table_frame, orient=VERTICAL)
    table = ttk.Treeview(table_frame,
                         columns=table_list[0],
                         show='headings',
                         xscrollcommand=xscroll.set,
                         yscrollcommand=yscroll.set
                         )

    # define column， not display
    for field_id in table_list[0]:
        table.column(field_id)

    # display column identifiers
    for field_id in table_list[0]:
        table.heading(field_id, text=field_id)

    # scroll bar
    xscroll.config(command=table.xview)
    xscroll.pack(side=BOTTOM, fill=X)
    yscroll.config(command=table.yview)
    yscroll.pack(side=RIGHT, fill=Y)
    table.pack(fill=BOTH, expand=True)

    table_menu = Menu(root_table)
    table_menu.add_command(label='Update',
                           font=("Times New Roman", 12),
                           command=lambda: insert_table(table, table_list[1::]))
    table_menu.add_command(label='Clean',
                           font=("Times New Roman", 12),
                           command=lambda: clean_table(table))
    table_menu.add_command(label='Save',
                           font=("Times New Roman", 12),
                           command=lambda: save_table_2_xlsx(table_list))
    table_menu.add_command(label='Quit',
                           font=("Times New Roman", 12),
                           command=lambda: quit_table(root_table))
    root_table.config(menu=table_menu)


def clean_table(table):
    for item in table.get_children():
        table.delete(item)


def save_table_2_xlsx(table_list):
    df = pd.DataFrame(columns=table_list[0], data=table_list[1::])
    out_xlsx = asksaveasfilename(filetypes=[("XLSX file", ".xlsx")],
                                 defaultextension=".xlsx")
    df.to_excel(out_xlsx, index=False)


def quit_table(root_table):
    root_table.destroy()


def insert_table(table, table_lst):
    for info_lst in table_lst:
        # print(info_lst)
        table.insert('', END, values=info_lst)  # 添加数据到末尾


def quit_app():
    root.destroy()


def clear_searchbox():
    search_box.delete(1.0, 'end')


def rm_duplicate():
    # obtail content in scrolledTExt box
    entries_str = search_box.get(1.0, 'end')
    entries_lst = entries_str.split('\n')
    unique_lst = []
    for entry in entries_lst:
        if not entry:
            continue
        if entry not in unique_lst:
            unique_lst.append(entry)

    search_box.delete(1.0, 'end')
    for entry in unique_lst:
        search_box.insert('end', entry)
        search_box.insert('end', '\n')


number_of_indexfungorum_records = StringVar(mainframe, value='0 record, 0 page')


def search_indexfungorum():
    # Confirm the term is correct
    entries_lst = search_box.get(1.0, 'end').split('\n')
    current_entry = current_entry_text.get()
    if not current_entry:
        messagebox.showerror("Error",
                             "please input searching term in the left input box")
        return

    if current_entry not in entries_lst:
        messagebox.showerror('Error',
                             f'{current_entry} not in the left input box')
        return

    # Normally, searching term is genus name. we later use it as the output prefix.
    genus_name = current_entry.strip()

    # Tell the path of chromedriver.exe
    driver_path = root_path + '\\\\' + 'chromedriver.exe'

    # Obtain number of pages and record number
    # activate chrome
    chrom_service = Service(driver_path)
    print(driver_path)
    driver = webdriver.Chrome(service=chrom_service)

    driver.implicitly_wait(5)

    # indexfungorum mainpage
    main_url = 'http://www.indexfungorum.org'

    # open mainpage
    driver.get(main_url)
    driver.set_window_position(700, 400, windowHandle='current')
    # set driver windows position

    # open search page
    search_index_fungorum_element = driver.find_element(By.CSS_SELECTOR, '[href="./Names/Names.asp"]')
    time.sleep(5)
    search_index_fungorum_element.click()

    # start searching
    driver.find_element(By.NAME, 'SearchTerm').send_keys(genus_name)
    driver.find_element(By.CSS_SELECTOR, '[type="submit"]').click()
    time.sleep(10)

    # obtain the number of records
    number_of_records = re.search(r'>of (\d+) records', driver.page_source, re.S).group(1)
    num_page = int(number_of_records) // 200 + 1

    # show number of records
    number_of_indexfungorum_records.set(f'{number_of_records} records, {num_page} pages')

    # get URLs of all records
    elements = driver.find_elements(By.CSS_SELECTOR, '[class="LinkColour1"]')
    record_url_lst = [element.get_attribute('href') for element in elements]
    time.sleep(5)

    if num_page >= 2:
        for i in range(2, num_page + 1):
            css_selector = f'[href="Names.asp?pg={i}"]'
            element = driver.find_element(By.CSS_SELECTOR, css_selector)

            element.click()
            time.sleep(5)
            elements = driver.find_elements(By.CSS_SELECTOR, '[class="LinkColour1"]')
            record_url_lst.extend([element.get_attribute('href') for element in elements])
    # DUBUG
    # print(f'No. URLs of records: {len(record_url_lst)}')
    # print(record_url_lst)

    # obtain table
    table_list = []
    head_lst = ["Record_name",
                'lineage',
                "Url",
                "Record_year",
                "Record_reference",
                "Current_name",
                "Current_name_reference",
                "Typification",
                "HostSubstratum",
                "Locality",
                "IndexFungorumID"]
    table_list.append(head_lst)

    # parse the page of each record
    num_url = 0
    for url in record_url_lst:
        num_url += 1
        driver.get(url)
        element = driver.find_element(By.XPATH, '//h3/..')
        details = element.get_attribute('outerHTML')

        # split content by paragraph
        # strip italic, bold and link signs
        content_lst = [re.sub('</p>|<b>|</b>|<i>|</i>|</a>', '', line)
                       for line in details.split('<p>')]

        content_lst2 = [line for line in content_lst
                        if "Instance" not in line and "Please contact" not in line]

        # Initial values
        record_name = 'NA'
        record_reference = 'NA'
        record_year = 'NA'
        lineage = 'NA'
        record_current_name = 'NA'
        record_current_name_reference = 'NA'
        index_fungorum_id = 'NA'
        record_typification = 'NA'
        host_substratum = 'NA'
        locality = 'NA'

        for line in content_lst2:
            # print(line)
            if "Names.asp?strGenus=" in line:
                record_name = ' '.join(line.split('>')[1].split(' ')[:2])
                record_reference = ' '.join(line.split('>')[1].split(' ')[2:])
                try:
                    record_year = re.search(' \(\d{4}\)', line, re.S).group(0).replace('(', '').replace(')', '')
                except AttributeError:
                    record_year = 'NA'
                # print(record_name.strip(), record_reference.strip(), record_year.strip(), sep = '\n')

            if "Position in classification" in line:
                lineage = line.split('<br>')[1].strip()
                # print(lineage)

            if "Species Fungorum current name" in line:
                record_current_name_info = line.split('">')[1]
                record_current_name = ' '.join(record_current_name_info.split(' ')[:2])
                record_current_name_reference = ' '.join(record_current_name_info.split(' ')[2:])
                # print(record_current_name,record_current_name_reference,sep ='\n')

            if "Index Fungorum Registration Identifier" in line:
                index_fungorum_id = re.search(r'\d{6}', line, re.S).group(0)
                # print(indexFungorumID)

            if "Typification Details" in line:
                record_typification = re.sub(r'<br>|</br>', '', line).split(':')[1].strip()
                # print(record_typification)

            if "Host-Substratum/Locality" in line:
                try:
                    if ":" in line:
                        host_substratum, locality = line.split('<br>')[1].split(':')
                        locality = locality.strip()
                    else:
                        host_substratum = line.split('<br>')[1].strip()
                except ValueError:
                    pass
                    # print(url, file=sys.stderr, flush=True)
                    # print(line, file=sys.stderr, flush=True)

                # print(HostSubstratum, locality)

        out_line_lst = [record_name,
                        lineage,
                        url,
                        record_year,
                        record_reference,
                        record_current_name,
                        record_current_name_reference,
                        record_typification,
                        host_substratum,
                        locality,
                        index_fungorum_id]
        # print(out_line_lst)
        table_list.append(out_line_lst)
    driver.close()

    # put data into Frame2
    table_window(table_list)


def click_last():
    entries_lst = search_box.get(1.0, 'end').split('\n')
    current_entry = current_entry_text.get()
    if not current_entry:
        messagebox.showerror("Error",
                             "please input searching term in the left input box")
    if current_entry not in entries_lst:
        messagebox.showerror('Error',
                             f'{current_entry} not in the left input box')
    index = entries_lst.index(current_entry) - 1
    print(index)
    current_entry = entries_lst[index]
    current_entry_text.delete(0, 'end')
    current_entry_text.insert(0, current_entry)


def click_next():
    entries_lst = search_box.get(1.0, 'end').split('\n')
    current_entry = current_entry_text.get()
    if not current_entry:
        index = 0
    else:
        index = entries_lst.index(current_entry) + 1
    if current_entry not in entries_lst:
        messagebox.showerror('Error',
                             f'{current_entry} not in the left input box')
    current_entry = entries_lst[index]
    current_entry_text.delete(0, 'end')
    current_entry_text.insert(0, current_entry)


# 1 label
input_lab = ttk.Label(mainframe, text="Enter searching term(s):",
                      font=("Times New Roman", 12, 'bold'))
input_lab.grid(row=0, column=0,
               padx=5, pady=3,
               sticky=W)

# 1 input box
search_box = scrolledtext.ScrolledText(mainframe,
                                       bd=2,
                                       wrap=WORD,
                                       font=("Times New Roman", 12),
                                       width=80, height=1,
                                       )
search_box.grid(row=1, column=0,
                rowspan=8,
                padx=5,
                sticky=E + W + S + N)
search_box.focus()

# 1 button
unique_button = Button(mainframe,
                       text='Remove duplicate', fg='blue',
                       font=("Times New Roman", 12),
                       width=10, height=1,
                       command=rm_duplicate)
unique_button.grid(row=2, column=1,
                   padx=5, pady=3,
                   ipadx=10,
                   sticky=W)
# 2 button
clear_button = Button(mainframe,
                      text='Clear', fg='blue',
                      font=("Times New Roman", 12),
                      width=10, height=1,
                      command=clear_searchbox)
clear_button.grid(row=2, column=2,
                  padx=5, pady=3,
                  ipadx=10,
                  sticky=W)

# 2 label
current_entry_lab = ttk.Label(mainframe,
                              text="Current term:",
                              font=("Times New Roman", 12, 'bold'))

current_entry_lab.grid(row=1, column=1,
                       padx=5, pady=3,
                       ipadx=10,
                       sticky=W)

# 2 input box
current_entry_text = Entry(mainframe, width=20,
                           font=("Times New Roman", 12))
current_entry_text.grid(row=1, column=2,
                        padx=5, pady=3)

unique_entry_list = search_box.get(1.0, 'end').split('\t')

# 3 button
next_button = Button(mainframe,
                     text='Next',
                     font=("Times New Roman", 12),
                     fg='blue',
                     width=10, height=1,
                     command=click_next)
next_button.grid(row=3, column=1,
                 padx=5, pady=3,
                 ipadx=10,
                 sticky=W)

# 4 button
last_button = Button(mainframe,
                     text='Last',
                     font=("Times New Roman", 12),
                     fg='blue',
                     width=10, height=1,
                     command=click_last)
last_button.grid(row=3, column=2,
                 padx=5, pady=3,
                 ipadx=10,

                 sticky=W)

# 5 button
search_indexfungorum_button = Button(mainframe,
                                     text='Search Index Fungorum',
                                     font=("Times New Roman", 12),
                                     fg='blue',
                                     width=15, height=1,
                                     command=search_indexfungorum)
search_indexfungorum_button.grid(row=4, column=1,
                                 padx=5, pady=3,
                                 ipadx=10,
                                 sticky=W)

# 6 button
search_mycobank_button = Button(mainframe,
                                text='Search MycoBank',
                                font=("Times New Roman", 12),
                                fg='blue',
                                width=15, height=1,
                                command=search_indexfungorum)
search_mycobank_button.grid(row=5, column=1,
                            padx=5, pady=3,
                            ipadx=10,
                            sticky=W)

# 3 label
number_indexfungorum_lab = Label(mainframe,
                                 textvariable=number_of_indexfungorum_records,
                                 font=("Times New Roman", 14, 'bold'),
                                 fg='red')
number_indexfungorum_lab.grid(row=6, column=1,
                              padx=10, pady=3,
                              columnspan=2)

# 7 button
quit_button = Button(mainframe,
                     text='Exit',
                     font=("Times New Roman", 12),
                     command=quit_app)
quit_button.grid(row=7,
                 column=1,
                 padx=5,
                 pady=3,
                 ipadx=10,
                 sticky=W)

root.mainloop()
