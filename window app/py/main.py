import subprocess
try:
    import tkinter as tk
except ImportError:
    try:
        subprocess.run(["pip", "install", "tkinter"])
        import tkinter as tk  # Try importing again
    except Exception as install_error:
        print(f"Error installing Tkinter: {install_error}")

import time
import json
import base64
import tempfile
try:
    import requests
except ImportError:
    try:
        subprocess.run(["pip", "install", "requests"])
        import requests  # Try importing again
    except Exception as install_error:
        print(f"Error installing Tkinter: {install_error}")
try:
    import win32api
except ImportError:
    try:
        subprocess.run(["pip", "install", "pywin32"])
    except:
        print('')
    try:
        import win32api
        subprocess.run(["pip", "install", "win32api"])
        import win32api  # Try importing again
    except Exception as install_error:
        print(f"Error installing Tkinter: {install_error}")
try:
    import win32print
except ImportError:
    try:
        import subprocess
        subprocess.run(["pip", "install", "win32print"])
        import win32print  # Try importing again
    except Exception as install_error:
        print(f"Error installing Tkinter: {install_error}")

class CustomApp:
    def __init__(self, master):
        self.master = master
        master.title("Printer Project")

        self.start_button = tk.Button(master, text="Start", command=self.start_process)
        self.start_button.grid(row=0, column=0, padx=5, pady=5)

        self.stop_button = tk.Button(master, text="Stop", command=self.stop_process)
        self.stop_button.grid(row=0, column=1, padx=5, pady=5)

        self.run_now_button = tk.Button(master, text="Run Now", command=self.run_now)
        self.run_now_button.grid(row=0, column=2, padx=5, pady=5)

        self.item_listbox = tk.Listbox(master)
        self.item_listbox.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

        self.is_running = False
        self.loop_task = None
        self.value = '0db609df72378bb1a30b320363dc93f80e9bf611'
        self.key = 'afb5bffb108ed229293a2f817ecdff739cd58acc'
        # self.start_process()

    def start_process(self):
        if not self.is_running:
            self.is_running = True
            self.loop_task = self.master.after(0, self.infinite_loop)

    def stop_process(self):
        self.is_running = False
        if self.loop_task:
            self.master.after_cancel(self.loop_task)

    def run_now(self):
        response = requests.get(
            'https://workspace.garryandnathan.com/printer/report/get?key=%s&value=%s' % (self.key, self.value))
        data = response.json()
        done_list = []

        for d in data:
            try:
                filename = tempfile.mktemp(".pdf")
                open(filename, "wb").write(base64.b64decode(d['data']))
                print(d.get('id', None))
                # Commenting out win32api and win32print for Tkinter compatibility
                win32api.ShellExecute(0,"print",filename,'"%s"' % win32print.GetDefaultPrinter(),".",0)
                self.item_listbox.insert(tk.END, "ID %s , report %s" % (d.get('id', None), 'Name'))
                done_list.append(d.get('id', None))
            except Exception as e:
                print(str(e))

        if done_list:
            r = requests.get('https://workspace.garryandnathan.com/printer/report/set?id=%s&key=%s&value=%s' % (
            done_list, self.key, self.value))
            data = r.json()
            if data.get('result', None) == 'success':
                pass

    def infinite_loop(self):
        if self.is_running:
            self.run_now()
            # Schedule the next iteration after 10 minutes (600,000 milliseconds)
            self.loop_task = self.master.after(600000, self.infinite_loop)


if __name__ == "__main__":
    root = tk.Tk()
    app = CustomApp(root)
    root.mainloop()
