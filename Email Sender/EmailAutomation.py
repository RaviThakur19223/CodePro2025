import smtplib
import csv
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

stop_flag = False

def send_emails():
    sender_email = sender_email_var.get().strip()
    email_password = email_password_var.get().strip()
    csv_file_path = csv_file_var.get().strip()
    subject = subject_var.get().strip()
    message_template = message_template_var.get("1.0", tk.END).strip()

    EMAIL_LIMIT = 190

    if not sender_email or not email_password or not csv_file_path or not subject or not message_template:
        messagebox.showerror("Error", "All fields are required.")
        return

    def process_emails():
        global stop_flag
        try:
            with open(csv_file_path, 'r') as file:
                reader = csv.DictReader(file)

                if 'name' not in reader.fieldnames or 'email' not in reader.fieldnames:
                    messagebox.showerror("Error", "CSV file must contain 'name' and 'email' columns.")
                    return

                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(sender_email, email_password)

                email_count = 0

                for row in reader:
                    if stop_flag:
                        update_status("Email sending stopped.")
                        break

                    name = row['name']
                    recipient_email = row['email']

                    personalized_message = f"Subject: {subject}\n\n{message_template.format(name=name)}"

                    try:
                        server.sendmail(sender_email, recipient_email, personalized_message)
                        update_status(f"Email sent to {name} ({recipient_email})")
                        email_count += 1
                    except Exception as e:
                        update_status(f"Failed to send email to {name} ({recipient_email}): {e}")

                    if email_count >= EMAIL_LIMIT:
                        update_status("Email limit reached. Pausing for 1 hour...")
                        server.quit()
                        time.sleep(3600)
                        server = smtplib.SMTP("smtp.gmail.com", 587)
                        server.starttls()
                        server.login(sender_email, email_password)
                        email_count = 0

            if not stop_flag:
                update_status("All emails processed.")
        except FileNotFoundError:
            messagebox.showerror("Error", "The specified CSV file was not found.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            try:
                server.quit()
            except:
                pass

    threading.Thread(target=process_emails, daemon=True).start()

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    csv_file_var.set(file_path)

def stop_sending():
    global stop_flag
    stop_flag = True

def update_status(message):
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, message + '\n')
    status_text.see(tk.END)
    status_text.config(state=tk.DISABLED)

root = tk.Tk()
root.geometry("600x600")
root.resizable(False, False)
root.title("Email Sender - By Ravi Thakur")
root.configure(bg="lightgreen")

sender_email_var = tk.StringVar()
email_password_var = tk.StringVar()
csv_file_var = tk.StringVar()
subject_var = tk.StringVar()

frame = tk.Frame(root, padx=10, pady=10, bg="lightgreen")
frame.pack(fill=tk.BOTH, expand=True)

row_count = 0

def create_label_and_entry(frame, label_text, text_var, show=None):
    global row_count
    tk.Label(frame, text=label_text, bg="lightgreen", anchor="w", width=15).grid(row=row_count, column=0, padx=5, pady=5, sticky="w")
    tk.Entry(frame, textvariable=text_var, show=show, width=40).grid(row=row_count, column=1, padx=5, pady=5, sticky="w")
    row_count += 1

create_label_and_entry(frame, "Sender Email:", sender_email_var)
create_label_and_entry(frame, "Email Password:", email_password_var, show="*")
tk.Label(frame, text="CSV File Path:", bg="lightgreen", anchor="w", width=15).grid(row=row_count, column=0, padx=5, pady=5, sticky="w")
tk.Entry(frame, textvariable=csv_file_var, width=40).grid(row=row_count, column=1, padx=5, pady=5, sticky="w")
tk.Button(frame, text="Browse", command=browse_file, bg="green", fg="white").grid(row=row_count, column=2, padx=5, pady=5)
row_count += 1
create_label_and_entry(frame, "Subject:", subject_var)

tk.Label(frame, text="Message:", bg="lightgreen", anchor="w", width=15).grid(row=row_count, column=0, padx=5, pady=5, sticky="w")
message_template_var = tk.Text(frame, height=10, width=40)
message_template_var.grid(row=row_count, column=1, columnspan=2, padx=5, pady=5, sticky="w")
row_count += 1

send_button = tk.Button(frame, text="Send Emails", command=send_emails, bg="green", fg="white", width=20)
send_button.grid(row=row_count, column=1, pady=10)

row_count += 1
tk.Label(frame, text="Status:", bg="lightgreen", anchor="w", width=15).grid(row=row_count, column=0, padx=5, pady=5, sticky="w")
status_text = tk.Text(frame, height=10, width=50, state=tk.DISABLED)
status_text.grid(row=row_count, column=1, columnspan=2, padx=5, pady=5)

row_count += 1
stop_button = tk.Button(frame, text="Stop", command=stop_sending, bg="red", fg="white", width=20)
stop_button.grid(row=row_count, column=1, pady=10)

root.mainloop()
