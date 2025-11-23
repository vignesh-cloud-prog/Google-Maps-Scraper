"""
This module contain the code for frontend
"""

from .communicator import Communicator
import tkinter as tk
from tkinter import ttk, WORD, filedialog
from scraper.scraper import Backend
from .common  import Common
import threading
import pandas as pd
import os



class Frontend:
    def __init__(self):
        """Initializing frontend layout"""

        self.root = tk.Tk()
        icon = tk.PhotoImage(file="Google map scraper\images\GMS.png")

        self.root.iconphoto(True, icon)
        self.root.geometry("850x650")
        self.root.resizable(False, False)
        self.root.title("Google maps scraper")

        self.style = ttk.Style()
        self.style.map(
            "my.TButton",  # making style for our buttons
            foreground=[("active", "green")],
            background=[("active", "white")],
        )

        bgimage = tk.PhotoImage(file="Google map scraper\images\Home.png")

        self.imglabel = tk.Label(self.root, image=bgimage)
        self.imglabel.place(x=0, y=0, relwidth=1, relheight=1)
        self.imglabel.image = bgimage

        """For search entry"""
        self.search_label = ttk.Label(
            self.root,
            text="Search:",
            font=("Franklin Gothic Medium", 17),
            foreground="green",
            background="white",
        )
        self.search_label.place(x=255, y=175)

        self.search_box = ttk.Entry(self.root, width=30, font=("Arial", 15))
        self.search_box.place(x=355, y=180)

        """For submit button"""
        self.submit_button = ttk.Button(
            self.root,
            text="Submit",
            width=15,
            command=self.getinput,
            style="my.TButton",
        )
        self.submit_button.place(x=355, y=320)

        """for output format entry"""
        self.outputFormatButtonLabel = ttk.Label(
            self.root,
            text="Format:",
            font=("Franklin Gothic Medium", 17),
            foreground="green",
            background="white",
        )
        self.outputFormatButtonLabel.place(x=255, y=230)

        self.outputFormatButton = ttk.Combobox(
            self.root, values=["Excel", "Json", "Csv"], state="readonly"
        )
        self.outputFormatButton.place(x=355, y=240)

        """For Excel file upload"""
        self.upload_label = ttk.Label(
            self.root,
            text="Upload Excel:",
            font=("Franklin Gothic Medium", 17),
            foreground="green",
            background="white",
        )
        self.upload_label.place(x=255, y=280)

        self.upload_button = ttk.Button(
            self.root,
            text="Browse Excel File",
            width=15,
            command=self.upload_excel_file,
            style="my.TButton",
        )
        self.upload_button.place(x=355, y=285)

        self.uploaded_file_label = ttk.Label(
            self.root,
            text="No file selected",
            font=("Arial", 10),
            foreground="gray",
            background="white",
        )
        self.uploaded_file_label.place(x=500, y=290)

        self.uploaded_file_path = None
        self.queries_from_file = []
        self.is_batch_processing = False

        """for message box"""
        self.show_text = tk.Text(
            self.root,
            font=("arial", 13),
            height=10,
            width=35,
            state="disabled",
            border=False,
            wrap=WORD,
            highlightbackground="black",
            highlightthickness=2,
        )
        self.show_text.place(x=295, y=440)

        """For healdess checkbox"""

        self.healdessCheckBoxVar = tk.IntVar()
        self.healdessCheckBox = tk.Checkbutton(
            self.root, text="Headless mode", variable=self.healdessCheckBoxVar)
        self.healdessCheckBox.place(x=700, y=45)

        self.__replacingtext(
            "Welcome to Google Maps Scraper!\n\nLet's start scraping..."
        )

        self.init_communicator()
    
    def init_communicator(self):
        Communicator.set_frontend_object(self)

    def upload_excel_file(self):
        """Handle Excel file upload and extract queries"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.uploaded_file_path = file_path
            filename = os.path.basename(file_path)
            self.uploaded_file_label.config(text=filename[:30] + "..." if len(filename) > 30 else filename, foreground="green")
            
            try:
                # Read Excel file and extract queries
                df = pd.read_excel(file_path)
                
                # Check if 'query' column exists
                if 'query' not in df.columns:
                    self.__replacingtext("Error: Excel file must contain a 'query' column")
                    self.uploaded_file_path = None
                    self.uploaded_file_label.config(text="No file selected", foreground="gray")
                    return
                
                # Extract queries (remove NaN values)
                self.queries_from_file = df['query'].dropna().astype(str).tolist()
                
                if len(self.queries_from_file) == 0:
                    self.__replacingtext("Error: No valid queries found in the 'query' column")
                    self.uploaded_file_path = None
                    self.uploaded_file_label.config(text="No file selected", foreground="gray")
                else:
                    self.__replacingtext(f"Successfully loaded {len(self.queries_from_file)} queries from Excel file")
                    
            except Exception as e:
                self.__replacingtext(f"Error reading Excel file: {str(e)}")
                self.uploaded_file_path = None
                self.uploaded_file_label.config(text="No file selected", foreground="gray")
                self.queries_from_file = []

    def __replacingtext(self, text):
        """This function will insert the text in text showing box"""

        self.show_text.config(state="normal")
        self.show_text.insert(tk.END, "• " + text)
        self.show_text.insert(tk.END, "\n\n")
        self.show_text.see(tk.END)
        self.show_text.config(state="disabled")

    def getinput(self):
        self.outputFormatValue = self.outputFormatButton.get()

        if len(self.outputFormatValue) == 0:
            self.__replacingtext(text="Oops! You did not select output format")
            return

        self.outputFormatValue = self.outputFormatValue.lower()
        self.headlessMode = self.healdessCheckBoxVar.get()

        # Check if Excel file is uploaded
        if self.uploaded_file_path and len(self.queries_from_file) > 0:
            # Process multiple queries from Excel file
            self.is_batch_processing = True
            self.submit_button.config(state="disabled")
            self.upload_button.config(state="disabled")
            self.threadToStartBackend = threading.Thread(
                target=self.process_multiple_queries)
            self.threadToStartBackend.start()
        else:
            # Process single query from search box
            self.searchQuery = self.search_box.get()

            if len(self.searchQuery) == 0:
                self.__replacingtext(text="Oops! You did empty search. Either enter a query or upload an Excel file.")
                return

            self.submit_button.config(state="disabled")
            self.searchQuery = self.searchQuery.lower()

            self.threadToStartBackend = threading.Thread(
                target=self.startscraping)
            self.threadToStartBackend.start()

    def closingbrowser(self):
        """It will close the browser when the app is closed"""

        try:
            # super().closeThread.set()
            Common.set_close_thread()
            self.root.destroy()
        except:
            pass

    def startscraping(self):
        backend = Backend(
            self.searchQuery,
            self.outputFormatValue,
            healdessmode=self.headlessMode
        )

        backend.mainscraping()
    
    def process_multiple_queries(self):
        """Process multiple queries from Excel file sequentially"""
        total_queries = len(self.queries_from_file)
        self.__replacingtext(f"Starting batch processing of {total_queries} queries...")
        
        for index, query in enumerate(self.queries_from_file, 1):
            if Common.close_thread_is_set():
                self.__replacingtext("Processing cancelled by user")
                break
                
            query = str(query).strip()
            if not query:
                continue
                
            self.__replacingtext(f"Processing query {index}/{total_queries}: {query}")
            
            try:
                backend = Backend(
                    query.lower(),
                    self.outputFormatValue,
                    healdessmode=self.headlessMode
                )
                backend.mainscraping()
                
                self.__replacingtext(f"Completed query {index}/{total_queries}: {query}")
            except Exception as e:
                self.__replacingtext(f"Error processing query {index}/{total_queries} ({query}): {str(e)}")
                continue
        
        self.__replacingtext(f"Batch processing completed! Processed {total_queries} queries.")
        self.is_batch_processing = False
        self.end_processing()
    
    def end_processing(self):
        # Only re-enable buttons if not in batch processing mode
        # (during batch processing, buttons should stay disabled until all queries are done)
        if not self.is_batch_processing:
            self.submit_button.config(state="normal")
            self.upload_button.config(state="normal")
        try:
            if hasattr(self, 'threadToStartBackend') and self.threadToStartBackend.is_alive():
                self.threadToStartBackend.join()
        except:
            pass

    def messageshowing(
        self,
        message):

        self.__replacingtext(message)


if __name__ == "__main__":
    app = Frontend()
    app.root.protocol("WM_DELETE_WINDOW", app.closingbrowser)
    app.root.mainloop()
