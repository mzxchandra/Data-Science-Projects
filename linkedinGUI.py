import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import sqlite3
from datetime import datetime
import webbrowser
import pyperclip
import json
import os
import sys
from linkedindriver import recipient_fetch_script

# Helper functions to manage paths and ensure directories exist
def get_base_path():
    if getattr(sys, 'frozen', False):
        # If the application is frozen (i.e., running as an executable)
        return sys._MEIPASS  # This is the temporary folder created by PyInstaller
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_user_data_path():
    # Use a folder in the user's home directory for storing config and templates
    return os.path.join(os.path.expanduser("~"), "linkedin-outreach-data")

def get_config_path():
    return os.path.join(get_user_data_path(), '.config.json')

def get_templates_path():
    return os.path.join(get_user_data_path(), '.templates.json')

def get_recently_contacted_path():
    return os.path.join(get_user_data_path(), '.recently_contacted.json')

def ensure_user_data_path_exists() -> bool:
    os.makedirs(get_user_data_path(), exist_ok=True)


class LinkedInOutreachApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LinkedIn Outreach Helper MVP")

        # Ensure the user data path exists before doing anything else
        print(f"path exists: {ensure_user_data_path_exists()}")
        print(f"user path is {get_user_data_path()}")
        print(f"config path is {get_config_path()}")
        print(f"templates path is {get_templates_path()}")
        print(f"recently contacted path is {get_recently_contacted_path()}")

        # Initialize variables
        self.filepath = ""
        self.data = None
        self.current_lead = None
        self.message_var = tk.StringVar()
        self.strength_var = tk.StringVar(value="Moderate")
        self.password = None  # Initialize password as None

        # Connect to the database
        self.conn = sqlite3.connect(os.path.join(get_user_data_path(), 'linkedin_contacts.db'))
        self.c = self.conn.cursor()
        self.c.execute('''
            CREATE TABLE IF NOT EXISTS contacts (
                linkedin_url TEXT PRIMARY KEY,
                last_contacted DATE,
                message_used TEXT,
                sent TEXT DEFAULT 'No'
            )
        ''')
        self.conn.commit()

        # Default templates (keep a copy of these for reversion)
        self.default_templates = {
            "exec": {
                "Strong": "Hi {name}, we've been connected for a while, and I thought you might be interested in joining our B2B pilot program. Given your role as {role} at {company}, your insights would be invaluable.",
                "Moderate": "Hi {name}. We're launching a B2B pilot program, and I thought you might be a great fit to provide feedback in your leadership role. Would love to hear your thoughts!",
                "Weak": "Hi {name}, I noticed you're {role} at {company}. I wanted to share that our app beams is now on the appstore, and thought it might be a useful tool for you."
            },
            "non-exec": {
                "Strong": "Hi {name}, as someone who's been in my network for a while, I wanted to personally invite you to check out our new app. I think it could be a great fit for someone in your role at {company}.",
                "Moderate": "Hi {name}, how is it going? We're excited to share our new app, which I think could be very relevant for your work at {company}. Would love for you to check it out!",
                "Weak": "Hi {name}, I noticed you're {role} at {company}. I wanted to share that our app beams is now on the appstore, and thought it might be a useful tool for you."
            }
        }

        #load templates
        self.templates = self.load_templates()

        # If the templates file was empty, save the default templates
        if not self.templates:
            self.templates = self.default_templates.copy()
            self.save_templates()

        # Create GUI elements
        self.create_widgets()

    def create_widgets(self):
        # Create a Notebook widget for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        # Create two frames, one for each tab
        self.exec_frame = ttk.Frame(self.notebook)
        self.non_exec_frame = ttk.Frame(self.notebook)
        self.contacted_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.exec_frame, text="Exec Results")
        self.notebook.add(self.non_exec_frame, text="Non-Exec Results")
        self.notebook.add(self.contacted_frame, text="Contacted Profiles")

        # Create Treeviews for each tab
        self.exec_tree = ttk.Treeview(self.exec_frame, columns=("First Name", "Last Name", "Position", "Company", "Connected On", "URL", "Message"), show='headings')
        self.non_exec_tree = ttk.Treeview(self.non_exec_frame, columns=("First Name", "Last Name", "Position", "Company", "Connected On", "URL", "Message"), show='headings')
        self.contacted_tree = ttk.Treeview(self.contacted_frame, columns=("Linkedin URL", "Last Contacted", "Message", "Visited"), show='headings')

        for tree in [self.exec_tree, self.non_exec_tree]:
            tree.heading("First Name", text="First Name")
            tree.heading("Last Name", text="Last Name")
            tree.heading("Position", text="Position")
            tree.heading("Company", text="Company")
            tree.heading("Connected On", text="Connected On")
            tree.heading("URL", text="URL")
            tree.heading("Message", text="Message")
            tree.pack(expand=True, fill="both")

        self.contacted_tree.heading("Linkedin URL", text="Linkedin URL")
        self.contacted_tree.heading("Last Contacted", text="Last Contacted")
        self.contacted_tree.heading("Message", text="Message")
        self.contacted_tree.heading("Visited", text="Visited")
        self.contacted_tree.pack(expand=True, fill="both")

        # Create the Load CSV button
        self.load_button = tk.Button(self.root, text="Load CSV", command=self.load_file)
        self.load_button.pack(pady=10)

        # Create a frame to hold the buttons side by side
        run_refresh_frame = tk.Frame(self.root)
        run_refresh_frame.pack(pady=10)

        # Run button to execute the filtering script
        self.run_button = tk.Button(run_refresh_frame, text="Hide Messaged Connections (Linkedin)", command=self.hide_contacted)
        self.run_button.pack(side="left", padx=5)

        # Refresh button to check recently messaged connections
        refresh_button = tk.Button(run_refresh_frame, text="Refresh Recently Contacted", command=self.refresh_recently_contacted)
        refresh_button.pack(side="left", padx=5)

        # Create a button for moving entries between tabs
        self.move_button = tk.Button(self.root, text="Move to Other Category", command=self.move_to_other_category)
        self.move_button.pack(pady=10)

        # Load data when switching to the third tab
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # Relationship Strength Dropdown Menu
        self.strength_options = ["Strong", "Moderate", "Weak"]
        self.strength_combobox = ttk.Combobox(self.root, values=self.strength_options, state="readonly")
        self.strength_combobox.pack(pady=10)

        # Bind the event to update the message when the selection changes
        self.strength_combobox.bind("<<ComboboxSelected>>", self.update_message_based_on_strength)

        # Message preview
        self.message_label = tk.Label(self.root, text="Generated Message:")
        self.message_label.pack()
        self.message_entry = tk.Entry(self.root, textvariable=self.message_var, width=100)
        self.message_entry.pack(pady=10)

        # Open profile button
        self.open_button = tk.Button(self.root, text="Open Profile", command=self.open_profile)
        self.open_button.pack(pady=10)

        # Create a frame to hold the buttons side by side
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # Mark as Sent button
        self.visit_button = tk.Button(button_frame, text="Mark as Sent", command=self.mark_as_sent)
        self.visit_button.pack(side="left", padx=5)

        # Skip button
        self.skip_button = tk.Button(button_frame, text="Skip Profile", command=self.skip_profile)
        self.skip_button.pack(side="left", padx=5)

        # Settings button
        self.settings_button = tk.Button(self.root, text="Settings", command=self.open_settings)
        self.settings_button.pack(pady=10)
        
    def load_templates(self):
        ensure_user_data_path_exists()
        templates_path = get_templates_path()
        if os.path.exists(templates_path):
            try:
                with open(templates_path, "r") as f:
                    self.templates = json.load(f)
            except json.JSONDecodeError:
                # If there's an error in loading the existing file, revert to default
                print("Error in reading templates.json, reverting to default templates.")
                self.templates = self.default_templates.copy()
                self.save_templates()
        else:
            # If the file doesn’t exist, create it with default templates
            self.templates = self.default_templates.copy()
            self.save_templates()
        return self.templates

        
    def save_templates(self):
        ensure_user_data_path_exists()
        templates_path = get_templates_path()
        print(templates_path)
        with open(templates_path, "w") as f:
            json.dump(self.templates, f)

    def load_credentials(self):
        ensure_user_data_path_exists()
        config_path = get_config_path()

        if os.path.exists(config_path):
            with open(config_path, "r") as config_file:
                config = json.load(config_file)
                self.username_entry.insert(0, config.get("linkedin_username", ""))

    def save_credentials(self):
        ensure_user_data_path_exists()
        config_path = get_config_path()

        config = {"linkedin_username": self.username_entry.get()}
        with open(config_path, "w") as config_file:
            json.dump(config, config_file)

        # Prompt user for password each time, or handle manually
        self.password = simpledialog.askstring("Password", "Enter your LinkedIn password:", show="*")

    def save_templates_and_close(self, settings_window):
        self.save_credentials()
        self.templates["exec"]["Strong"] = self.exec_strong_text.get("1.0", tk.END).strip()
        self.templates["exec"]["Moderate"] = self.exec_moderate_text.get("1.0", tk.END).strip()
        self.templates["exec"]["Weak"] = self.exec_weak_text.get("1.0", tk.END).strip()
        self.templates["non-exec"]["Strong"] = self.non_exec_strong_text.get("1.0", tk.END).strip()
        self.templates["non-exec"]["Moderate"] = self.non_exec_moderate_text.get("1.0", tk.END).strip()
        self.templates["non-exec"]["Weak"] = self.non_exec_weak_text.get("1.0", tk.END).strip()
        self.save_templates()

        if settings_window.winfo_exists():
            settings_window.destroy()

    def revert_to_default(self):
        self.templates = self.default_templates.copy()
        self.save_templates()
        self.update_settings_window()
        

    def update_settings_window(self):
        # Update the text fields in the settings window with current templates
        self.exec_strong_text.delete("1.0", tk.END)
        self.exec_strong_text.insert(tk.END, self.templates["exec"]["Strong"])

        self.exec_moderate_text.delete("1.0", tk.END)
        self.exec_moderate_text.insert(tk.END, self.templates["exec"]["Moderate"])

        self.exec_weak_text.delete("1.0", tk.END)
        self.exec_weak_text.insert(tk.END, self.templates["exec"]["Weak"])

        self.non_exec_strong_text.delete("1.0", tk.END)
        self.non_exec_strong_text.insert(tk.END, self.templates["non-exec"]["Strong"])

        self.non_exec_moderate_text.delete("1.0", tk.END)
        self.non_exec_moderate_text.insert(tk.END, self.templates["non-exec"]["Moderate"])

        self.non_exec_weak_text.delete("1.0", tk.END)
        self.non_exec_weak_text.insert(tk.END, self.templates["non-exec"]["Weak"])

    def on_tree_select(self, event):
        # Get the selected item
        selected_item = self.tree.selection()[0]
        item = self.tree.item(selected_item)

        # Extract relevant information from the selected item
        role = item['values'][1]
        name = item['values'][0]
        company = item['values'][3]

        # Determine segment and relationship strength based on your logic
        segment = self.determine_segment(role)
        relationship_strength = self.calculate_relationship_strength(item['values'][3])

        # Pre-select the relationship strength in the dropdown
        self.strength_combobox.set(relationship_strength)

        # Generate the message based on the pre-selected strength
        message = self.generate_message(name, role, company, relationship_strength, segment)

        # Update the message preview with the generated message
        self.message_var.set(message)

    def open_profile(self):
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 0:  # Exec tab
            selected_tree = self.exec_tree
        else:  # Non-Exec tab
            selected_tree = self.non_exec_tree

        selected_item = selected_tree.selection()[0]
        item = selected_tree.item(selected_item)
        url = item['values'][5]
        message = self.message_var.get()  # Get the current text from the Generated Message field
        webbrowser.open(url)
        pyperclip.copy(message)
        messagebox.showinfo("Message Copied", "The message has been copied to your clipboard!")

    def mark_as_sent(self):
        # Determine which treeview (exec_tree or non_exec_tree) is currently active
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 0:  # Exec tab
            selected_tree = self.exec_tree
        else:  # Non-Exec tab
            selected_tree = self.non_exec_tree

        # Get the currently selected item
        selected_item = selected_tree.selection()[0]
        item = selected_tree.item(selected_item)

        # Extract the URL to mark as visited
        url = item['values'][5]  # Assuming URL is at index 5
        print(f"Marking URL as visited: {url}")  # Debugging statement
        message = self.message_var.get()  # Get the current text from the Generated Message field
        
        self.log_message_sent(url, message, True)

        # Remove the item from the TreeView
        selected_tree.delete(selected_item)

        # Show a confirmation message
        messagebox.showinfo("Profile Marked", "This profile has been marked as visited and removed from the list.")

    def skip_profile(self):
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 0:  # Exec tab
            selected_tree = self.exec_tree
        else:  # Non-Exec tab
            selected_tree = self.non_exec_tree

        # Save the profile details in the database
        selected_item = selected_tree.selection()[0]
        item = selected_tree.item(selected_item)
        url = item['values'][5]
        self.log_message_sent(url, "", False)
        # Skip the current profile (e.g., remove it from the list)
        selected_tree.delete(selected_item)

    def load_contacted_profiles(self):
        # Clear the treeview first
        for item in self.contacted_tree.get_children():
            self.contacted_tree.delete(item)

        # Query the database for all contacted profiles
        self.c.execute("SELECT linkedin_url, last_contacted, message_used, sent FROM contacts")
        contacted_profiles = self.c.fetchall()

        # Insert data into the contacted_tree
        for profile in contacted_profiles:
            linkedin_url, last_contacted, message, sent = profile
            self.contacted_tree.insert("", "end", values=(linkedin_url, last_contacted, message, sent))


    def on_tab_changed(self, event):
            selected_tab = event.widget.index("current")
            if selected_tab == 2:  # Third tab index
                self.load_contacted_profiles()

    def log_message_sent(self, url, message, sent: bool):
        last_contacted = datetime.now().strftime("%Y-%m-%d")
        self.c.execute("INSERT OR REPLACE INTO contacts (linkedin_url, last_contacted, message_used, sent) VALUES (?, ?, ?, ?)", 
                (url, last_contacted, message, sent))
        self.conn.commit()

    def open_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")

        # Add fields for LinkedIn username and password
        tk.Label(settings_window, text="LinkedIn Username:").pack()
        self.username_entry = tk.Entry(settings_window, width=50)
        self.username_entry.pack()

        # Load existing credentials if available
        self.load_credentials()

        # Executive Strong template
        tk.Label(settings_window, text="Executive (Strong Relationship):").pack()
        self.exec_strong_text = tk.Text(settings_window, height=5, width=50)
        self.exec_strong_text.insert(tk.END, self.templates["exec"]["Strong"])
        self.exec_strong_text.pack()

        # Executive Moderate template
        tk.Label(settings_window, text="Executive (Moderate Relationship):").pack()
        self.exec_moderate_text = tk.Text(settings_window, height=5, width=50)
        self.exec_moderate_text.insert(tk.END, self.templates["exec"]["Moderate"])
        self.exec_moderate_text.pack()

        # Executive Weak template
        tk.Label(settings_window, text="Executive (Weak Relationship):").pack()
        self.exec_weak_text = tk.Text(settings_window, height=5, width=50)
        self.exec_weak_text.insert(tk.END, self.templates["exec"]["Weak"])
        self.exec_weak_text.pack()

        # Non-Executive Strong template
        tk.Label(settings_window, text="Non-Executive (Strong Relationship):").pack()
        self.non_exec_strong_text = tk.Text(settings_window, height=5, width=50)
        self.non_exec_strong_text.insert(tk.END, self.templates["non-exec"]["Strong"])
        self.non_exec_strong_text.pack()        

        # Non-Executive Moderate template
        tk.Label(settings_window, text="Non-Executive (Moderate Relationship):").pack()
        self.non_exec_moderate_text = tk.Text(settings_window, height=5, width=50)
        self.non_exec_moderate_text.insert(tk.END, self.templates["non-exec"]["Moderate"])
        self.non_exec_moderate_text.pack()

        # Non-Executive Weak template
        tk.Label(settings_window, text="Non-Executive (Weak Relationship):").pack()
        self.non_exec_weak_text = tk.Text(settings_window, height=5, width=50)
        self.non_exec_weak_text.insert(tk.END, self.templates["non-exec"]["Weak"])
        self.non_exec_weak_text.pack()

        # Save and Revert buttons
        save_button = tk.Button(settings_window, text="Save", command=lambda: self.save_templates_and_close(settings_window))
        save_button.pack(pady=10)

        revert_button = tk.Button(settings_window, text="Revert to Default", command=self.revert_to_default)
        revert_button.pack(pady=10)

    def generate_message(self, name, role, company, strength, segment):
    # Use the templates from the settings
        return self.templates[segment][strength].format(name=name, role=role, company=company)

    def update_message_based_on_strength(self, event):
        # Determine which treeview (exec_tree or non_exec_tree) is currently active
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 0:  # Exec tab
            selected_tree = self.exec_tree
        else:  # Non-Exec tab
            selected_tree = self.non_exec_tree

        # Get the selected relationship strength from the combobox
        selected_strength = self.strength_combobox.get()

        # Get the currently selected row in the tree
        selected_item = selected_tree.selection()[0]
        item = selected_tree.item(selected_item)

        # Extract the necessary data to regenerate the message
        name = item['values'][0]
        role = item['values'][2]  # Position is at index 2
        company = item['values'][3]  # Company is at index 3
        segment = self.determine_segment(role)

        # Generate the new message based on the newly selected strength
        new_message = self.generate_message(name, role, company, selected_strength, segment)
        
        # Update the message preview
        self.message_var.set(new_message)

    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if self.filepath:
            self.data = pd.read_csv(self.filepath)
            self.populate_table()

    def populate_table(self):
        for index, row in self.data.iterrows():
            # Check if the URL is already in the database
            self.c.execute("SELECT 1 FROM contacts WHERE linkedin_url=?", (row['URL'],))
            result = self.c.fetchone()

            if result:
                continue  # Skip this profile if it is already in the database

            segment = self.determine_segment(row['Position'])
            relationship_strength = self.calculate_relationship_strength(row['Connected On'])
            message = self.generate_message(row['First Name'], row['Position'], row['Company'], relationship_strength, segment)
            
            # Insert into the correct treeview based on the segment
            if segment == 'exec':
                self.exec_tree.insert("", "end", values=(row['First Name'],row['Last Name'],  row['Position'], row['Company'], row['Connected On'], row['URL'], message))
            else:
                self.non_exec_tree.insert("", "end", values=(row['First Name'], row['Last Name'], row['Position'], row['Company'], row['Connected On'], row['URL'], message))

    def determine_segment(self, role):
        # Convert role to string and handle NaN values
        if pd.isna(role):
            role = ""  # Treat NaN as an empty string
        else:
            role = str(role)

        exec_roles = ['CTO', 'CEO', 'Team Lead', 'Manager', 'Director', 'VP', 'co-founder', 'coo', 'cco']

        if any(title.lower() in role.lower() for title in exec_roles):
            return 'exec'
        else:
            return 'non-exec'

    def move_to_other_category(self):
        # Determine which tab is currently selected
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 0:  # Exec tab
            selected_tree = self.exec_tree
            target_tree = self.non_exec_tree
        else:  # Non-Exec tab
            selected_tree = self.non_exec_tree
            target_tree = self.exec_tree

        selected_item = selected_tree.selection()[0]
        item = selected_tree.item(selected_item)
        selected_tree.delete(selected_item)

        # Insert the item into the other treeview
        target_tree.insert("", "end", values=item['values'])

    def calculate_relationship_strength(self, connection_date):
        # Adjust the date format to match "07 Feb 2023"
        connection_date = datetime.strptime(connection_date, "%d %b %Y")
        duration = (datetime.now() - connection_date).days / 365
        if duration > 3:
            return "Strong"
        elif 1 <= duration <= 3:
            return "Moderate"
        else:
            return "Weak"

    def hide_contacted(self):
        json_file_path = get_recently_contacted_path()
        
        if not os.path.exists(json_file_path):
            self.refresh_recently_contacted()
            with open(json_file_path, 'r') as json_file:
                recently_contacted = json.load(json_file)
        else:
            with open(json_file_path, 'r') as json_file:
                recently_contacted = json.load(json_file)
        
        recently_contacted_split = []
        for full_name in recently_contacted:
            names = full_name.strip().lower().split()
            if len(names) >= 2:
                first_name, last_name = names[0], names[-1]
                recently_contacted_split.append((first_name, last_name))
            else:
                continue

        for tree in [self.exec_tree, self.non_exec_tree]:
            for item in tree.get_children():
                item_values = tree.item(item, "values")
                first_name_in_tree = item_values[0].strip().lower()
                last_name_in_tree = item_values[1].strip().lower()

                for first_name, last_name in recently_contacted_split:
                    if first_name_in_tree == first_name and last_name_in_tree == last_name:
                        print(f"Deleting: {first_name_in_tree} {last_name_in_tree} from {tree}")
                        tree.delete(item)
                        break

        tree.update_idletasks()

    def refresh_recently_contacted(self):
        ensure_user_data_path_exists()
        json_file_path = get_recently_contacted_path()
        recently_contacted = self.check_recently_contacted()
        with open(json_file_path, 'w') as json_file:
            json.dump(recently_contacted, json_file)
    
    def check_recently_contacted(self) -> list:
        # Load username from config file
        config_path = get_config_path()
        if os.path.exists(config_path):
            with open(config_path, "r") as config_file:
                config = json.load(config_file)
                username = config.get("linkedin_username")

        # Prompt user for password if it hasn't been set
        if not self.password:
            self.password = simpledialog.askstring("Password", "Enter your LinkedIn password:", show="*")

        # Ensure that username and password are retrieved correctly
        if not username or not self.password:
            messagebox.showerror("Error", "Username or Password not found. Please set them in the settings.")
            return []

        # Use the provided username and password to fetch recently contacted names
        try:
            recently_contacted = recipient_fetch_script(username, self.password)
            return recently_contacted
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch recently contacted profiles: {str(e)}")
            return []

           
# Initialize and run the application
root = tk.Tk()
app = LinkedInOutreachApp(root)
root.mainloop()