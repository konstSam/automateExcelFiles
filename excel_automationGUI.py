import openpyxl
import os
from datetime import datetime
from tkinter import Tk, filedialog, Label, Entry, Button, Text, END, W, E, N, S, messagebox, StringVar, Scrollbar


def update_percentages(folder_path, date_input, monthly_profit, profit, date, feedback):
    feedback.configure(state='normal')

    for item in os.listdir(folder_path):
        # Construct the full path to the item
        item_path = os.path.join(folder_path, item)

        # If the item is a directory, recursively call the function on the directory
        if os.path.isdir(item_path):
            update_percentages(item_path, date_input,
                               monthly_profit, profit, date, feedback)

        # If the item is a file and ends with ".xlsx"
        elif os.path.isfile(item_path) and item.endswith('.xlsx'):
            # Initialize a variable to keep track of whether a match was found
            match_found = False
            try:
                # Load the Excel file and select the worksheet you want to search
                workbook = openpyxl.load_workbook(item_path)
                worksheet = workbook.active

                # Get the name of the file
                filename = os.path.basename(item_path)

                # Search for the date in column A and update the corresponding cell in column B
                for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=2):
                    if row[0].value == date_input:
                        row[1].value = monthly_profit
                        print(
                            f'File {filename}: Added {profit}% monthly profit in {date}.')
                        match_found = True
                        break

            except:
                feedback.insert(
                    END, f"Error updating {filename}\n", 'error')

            # If no match was found, print a message
            if not match_found:
                feedback.insert(END, f"No match found in {filename}\n")
                workbook.close()
                continue

            # Save the changes to the Excel file
            try:
                workbook.save(os.path.join(folder_path, filename))
                feedback.insert(
                    END, f"Updated month {date} to {profit}% in {filename} and saved successfully.\n", 'updated')
            except PermissionError as e:
                workbook.close()
                feedback.insert(f"Error: {e}")

        else:
            feedback.insert(END, f"Invalid file {item_path}\n", 'error')


def finish_up(feedback):
    # completed update
    feedback.configure(state='normal')
    feedback.insert(END, f"Update completed!")
    messagebox.showinfo("Success", f"Update completed!")
    feedback.configure(state='disabled')


def validate_date(date_str):    # Date & validation and check function
    try:
        datetime.strptime(date_str, '%d/%m/%Y')
        return True

    except ValueError:
        return False


def check_date(event, month_entry, update_btn):
    if not validate_date(month_entry.get()):
        month_entry.config(fg='red')
        update_btn.config(state='disabled')
    else:
        month_entry.config(fg='green')
        update_btn.config(state='normal')


def get_date(month_entry):
    date = month_entry.get()
    if not validate_date(date):
        messagebox.showerror(
            "Error", ("Invalid date format!"))
        return
    else:
        return datetime.strptime(date, '%d/%m/%Y')


def update_files(folder_path, date_input, monthly_profit, profit, date, feedback, update_btn, get_parent_directory):
    # Update function
    if not folder_path:
        messagebox.showerror("Error", "Please select a folder first")
        return

    update_btn.config(state='disabled')
    update_percentages(folder_path, date_input,
                       monthly_profit, profit, date, feedback)
    finish_up(feedback)
    update_btn.config(state='normal')


########################################################################


def main():
    root = Tk()
    root.title("Excel Files Updater")

    # Set the geometry of the window
    root.geometry("865x510")
    root.resizable(False, False)
    folder_path = StringVar()

    # Input widgets
    month_label = Label(root, text="Enter month:", font=('Arial', 12, 'bold'))
    month_entry = Entry(root, font=('Arial', 12, 'bold'))
    month_label.grid(row=0, column=0)
    month_entry.grid(row=0, column=1)

    percent_label = Label(root, text="Enter percentage:",
                          font=('Arial', 12, 'bold'))
    percent_entry = Entry(root, font=('Arial', 12, 'bold'))
    percent_label.grid(row=1, column=0)
    percent_entry.grid(row=1, column=1)

    def get_parent_directory():
        selected_directory = filedialog.askdirectory()
        folder_path.set(selected_directory)

    choose_directory = Button(root, text="Choose Directory:", command=get_parent_directory,
                              background='dark grey', font=('Arial', 12, 'bold'))
    choose_directory.grid(row=2, column=0, columnspan=4)

    # Feedback text box
    feedback = Text(root, width=105)
    feedback.config(state='disabled')
    feedback.grid(row=3, column=0, columnspan=4)

    # Add a scrollbar
    scrollbar = Scrollbar(root)
    scrollbar.grid(row=3, column=4, sticky='NS')

    # Configure the feedback_box to use the scrollbar
    feedback.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=feedback.yview)

    month_entry.bind(
        '<FocusOut>', lambda event: check_date(event, month_entry, update_btn))

    # Update button
    update_btn = Button(root, text="Update Files", command=lambda: update_files(folder_path.get(), get_date(month_entry), float(percent_entry.get(
    ).replace(',', '.')) / 100, float(percent_entry.get().replace(',', '.')), month_entry.get(), feedback, update_btn, get_parent_directory),
        background='light blue', font=('Arial', 12, 'bold'))
    update_btn.grid(row=4, column=0, columnspan=4)

    # Configure tag styles
    feedback.tag_config('updated', foreground='green')
    feedback.tag_config('error', foreground='red')

    # Launch GUI
    root.mainloop()


if __name__ == '__main__':
    main()
