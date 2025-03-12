import sys
import requests
import re
import time
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLineEdit, QPushButton, QTextEdit, 
                            QLabel, QProgressBar, QMessageBox, QFileDialog,
                            QComboBox, QRadioButton, QButtonGroup)
from PyQt6.QtCore import QThread, pyqtSignal

class EmailScraper(QThread):
    progress_update = pyqtSignal(str)
    finished = pyqtSignal(list)
    current_results = pyqtSignal(list)  # New signal for current results
    
    def __init__(self, query, api_key):
        super().__init__()
        self.query = query
        self.api_key = api_key
        self.email_sources = []
        self.processed_emails = set()  # To prevent duplicate emails
        self._is_running = True  # Flag to control the scraping process

    def stop_scraping(self):
        """Stops the scraping process gracefully"""
        self._is_running = False
        self.progress_update.emit("Stopping the search process...")

    def extract_emails(self, text, source_url):
        """Extracts email addresses from the given text and adds them to the list"""
        found_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
        new_count = 0
        for email in found_emails:
            if email.lower() not in self.processed_emails:  # Check if email is not already processed
                self.email_sources.append((email, source_url))
                self.processed_emails.add(email.lower())  # Mark email as processed
                new_count += 1
                # Emit current results after each new email found
                self.current_results.emit(self.email_sources)
        return new_count

    def fetch_page_content(self, url):
        """Fetches content from the specified URL"""
        try:
            response = requests.get(url, timeout=10)
            return response.text
        except Exception as e:
            self.progress_update.emit(f"Failed to fetch page content ({url}): {str(e)}")
            return None

    def run(self):
        url = f"https://serpapi.com/search?engine=google&q={self.query}&api_key={self.api_key}&num=100"
        page_number = 1
        
        while url and self._is_running:  # Check if scraping should continue
            try:
                self.progress_update.emit(f"Scanning page {page_number}...")
                response = requests.get(url)
                data = response.json()

                if "organic_results" in data:
                    for result in data["organic_results"]:
                        if not self._is_running:  # Check if should stop
                            break
                            
                        site_url = result.get("link", "No Source")
                        snippet_text = result.get("snippet", "")

                        # Extract emails from snippet
                        new_emails = self.extract_emails(snippet_text, site_url)
                        if new_emails > 0:
                            self.progress_update.emit(f"Found {new_emails} new emails from snippet: {site_url}")

                        # Scan page content
                        if site_url != "No Source":
                            self.progress_update.emit(f"Performing detailed page scan: {site_url}")
                            try:
                                # First try to get direct page content
                                page_content = self.fetch_page_content(site_url)
                                if page_content:
                                    new_emails = self.extract_emails(page_content, site_url)
                                    if new_emails > 0:
                                        self.progress_update.emit(f"Found {new_emails} new emails from page content: {site_url}")
                                
                                # Try to get content through SerpAPI
                                page_response = requests.get(f"https://serpapi.com/search?engine=google&q={self.query}&api_key={self.api_key}&fetch_page={site_url}")
                                page_data = page_response.json()
                                new_emails = self.extract_emails(str(page_data), site_url)
                                if new_emails > 0:
                                    self.progress_update.emit(f"Found {new_emails} new emails from SerpAPI content: {site_url}")

                            except Exception as e:
                                self.progress_update.emit(f"Error during page scan: {str(e)}")
                                continue

                self.progress_update.emit(f"Total {len(self.email_sources)} unique emails found. Searching for next page...")
                url = data.get("serpapi_pagination", {}).get("next", None) if self._is_running else None
                page_number += 1
                time.sleep(2)  # Wait for rate limiting

            except Exception as e:
                self.progress_update.emit(f"Error occurred: {str(e)}")
                break

        status = "stopped" if not self._is_running else "completed"
        self.progress_update.emit(f"Search {status}! Total {len(self.email_sources)} unique email addresses found.")
        self.finished.emit(self.email_sources)

class BingEmailScraper(QThread):
    progress_update = pyqtSignal(str)
    finished = pyqtSignal(list)
    
    def __init__(self, query, api_key):
        super().__init__()
        self.query = query
        self.api_key = api_key
        self.email_sources = []
        self.processed_emails = set()  # To prevent duplicate emails
        self.headers = {
            'Ocp-Apim-Subscription-Key': self.api_key,
        }
        self.base_url = "https://api.bing.microsoft.com/v7.0/search"

    def extract_emails(self, text, source_url):
        """Extracts email addresses from the given text and adds them to the list"""
        found_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
        new_count = 0
        for email in found_emails:
            if email.lower() not in self.processed_emails:
                self.email_sources.append((email, source_url))
                self.processed_emails.add(email.lower())
                new_count += 1
        return new_count

    def fetch_page_content(self, url):
        """Fetches content from the specified URL"""
        try:
            response = requests.get(url, timeout=10)
            return response.text
        except Exception as e:
            self.progress_update.emit(f"Failed to fetch page content ({url}): {str(e)}")
            return None

    def run(self):
        offset = 0
        total_results = 50  # Bing API default limit per query
        
        while offset < total_results:
            try:
                params = {
                    'q': self.query,
                    'count': 50,  # Maximum results per page
                    'offset': offset,
                    'textDecorations': True,
                    'textFormat': 'HTML'
                }
                
                self.progress_update.emit(f"Scanning Bing results (offset: {offset})...")
                response = requests.get(self.base_url, headers=self.headers, params=params)
                
                if response.status_code != 200:
                    self.progress_update.emit(f"Error: Bing API returned status code {response.status_code}")
                    break
                
                data = response.json()
                
                if 'webPages' in data and 'value' in data['webPages']:
                    results = data['webPages']['value']
                    total_results = min(data['webPages'].get('totalEstimatedMatches', 50), 200)  # Limit to 200 results
                    
                    for result in results:
                        site_url = result.get('url', "No Source")
                        snippet_text = result.get('snippet', '')
                        
                        # Extract emails from snippet
                        new_emails = self.extract_emails(snippet_text, site_url)
                        if new_emails > 0:
                            self.progress_update.emit(f"Found {new_emails} new emails from snippet: {site_url}")
                        
                        # Scan page content
                        if site_url != "No Source":
                            self.progress_update.emit(f"Performing detailed page scan: {site_url}")
                            try:
                                page_content = self.fetch_page_content(site_url)
                                if page_content:
                                    new_emails = self.extract_emails(page_content, site_url)
                                    if new_emails > 0:
                                        self.progress_update.emit(f"Found {new_emails} new emails from page content: {site_url}")
                            
                            except Exception as e:
                                self.progress_update.emit(f"Error during page scan: {str(e)}")
                                continue
                    
                    offset += len(results)
                    self.progress_update.emit(f"Total {len(self.email_sources)} unique emails found. Moving to next page...")
                    time.sleep(0.5)  # Respect rate limits
                else:
                    break
                
            except Exception as e:
                self.progress_update.emit(f"Error occurred: {str(e)}")
                break
        
        self.progress_update.emit(f"Scan completed! Total {len(self.email_sources)} unique email addresses found.")
        self.finished.emit(self.email_sources)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Harvester")
        self.setMinimumSize(800, 600)
        self.setup_ui()
        self.current_results = []  # Store current results

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Search Engine Selection
        engine_layout = QHBoxLayout()
        engine_label = QLabel("Search Engine:")
        self.engine_group = QButtonGroup()
        self.google_radio = QRadioButton("Serpapi Google")
        self.bing_radio = QRadioButton("Bing")
        self.google_radio.setChecked(True)
        self.engine_group.addButton(self.google_radio)
        self.engine_group.addButton(self.bing_radio)
        engine_layout.addWidget(engine_label)
        engine_layout.addWidget(self.google_radio)
        engine_layout.addWidget(self.bing_radio)
        layout.addLayout(engine_layout)

        # API Key area
        api_key_layout = QHBoxLayout()
        self.api_key_label = QLabel("Serpapi Google API Key:")
        self.api_key_input = QLineEdit()
        self.api_key_input.setText("")
        self.api_key_input.setPlaceholderText("Enter your serpapi API key")
        api_key_layout.addWidget(self.api_key_label)
        api_key_layout.addWidget(self.api_key_input)
        layout.addLayout(api_key_layout)

        # Update API key label when search engine changes
        self.google_radio.toggled.connect(lambda: self.update_api_label("Serpapi Google"))
        self.bing_radio.toggled.connect(lambda: self.update_api_label("Bing"))

        # Search area with Stop button
        search_layout = QHBoxLayout()
        search_label = QLabel("Search Term:")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Example: "@example.com"')
        self.search_button = QPushButton("Search")
        self.stop_button = QPushButton("Stop")  # New Stop button
        self.stop_button.setEnabled(False)
        self.search_button.clicked.connect(self.start_search)
        self.stop_button.clicked.connect(self.stop_search)  # Connect Stop button
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)
        search_layout.addWidget(self.stop_button)
        layout.addLayout(search_layout)

        # Progress area
        self.progress_text = QTextEdit()
        self.progress_text.setReadOnly(True)
        layout.addWidget(self.progress_text)

        # Results area
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        layout.addWidget(self.results_text)

        # Save options
        save_options_layout = QHBoxLayout()
        
        # File format selection
        self.format_combo = QComboBox()
        self.format_combo.addItems(["Excel (.xlsx)", "CSV (.csv)", "Text (.txt)"])
        save_options_layout.addWidget(QLabel("File Format:"))
        save_options_layout.addWidget(self.format_combo)
        
        # Save button
        self.save_button = QPushButton("Save Results")
        self.save_button.clicked.connect(self.save_results)
        self.save_button.setEnabled(False)
        save_options_layout.addWidget(self.save_button)
        
        layout.addLayout(save_options_layout)

    def update_api_label(self, engine):
        self.api_key_label.setText(f"{engine} API Key:")
        if engine == "Google":
            self.api_key_input.setText("")
        else:
            self.api_key_input.setText("")  # Clear for Bing API key
            self.api_key_input.setPlaceholderText("Enter your Bing API key")

    def start_search(self):
        query = self.search_input.text()
        api_key = self.api_key_input.text()
        
        if not query or not api_key:
            QMessageBox.warning(self, "Warning", "Please enter API Key and search term!")
            return

        self.progress_text.clear()
        self.results_text.clear()
        self.current_results = []  # Reset current results
        self.save_button.setEnabled(True)  # Enable save button from start
        self.search_button.setEnabled(False)
        self.stop_button.setEnabled(True)  # Enable stop button

        # Create appropriate scraper based on selected search engine
        if self.google_radio.isChecked():
            self.scraper = EmailScraper(query, api_key)
        else:
            self.scraper = BingEmailScraper(query, api_key)

        self.scraper.progress_update.connect(self.update_progress)
        self.scraper.finished.connect(self.search_completed)
        self.scraper.current_results.connect(self.update_current_results)  # Connect to current results
        self.scraper.start()

    def stop_search(self):
        """Stops the current search process"""
        if hasattr(self, 'scraper'):
            self.scraper.stop_scraping()
            self.stop_button.setEnabled(False)
            self.search_button.setEnabled(True)

    def update_current_results(self, results):
        """Updates the current results in real-time"""
        self.current_results = results
        self.results_text.clear()
        for email, source in results:
            self.results_text.append(f"{email} - {source}")

    def update_progress(self, text):
        self.progress_text.append(text)

    def search_completed(self, results):
        self.current_results = results
        self.search_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)
        
        self.results_text.clear()
        for email, source in results:
            self.results_text.append(f"{email} - {source}")
        
        QMessageBox.information(self, "Information", f"Total {len(results)} email addresses found!")

    def save_results(self):
        try:
            # Convert current results to DataFrame
            df = pd.DataFrame(self.current_results, columns=['Email', 'Source'])
            
            # Determine file extension based on selected format
            format_type = self.format_combo.currentText()
            if "Excel" in format_type:
                file_filter = "Excel File (*.xlsx)"
                default_ext = ".xlsx"
            elif "CSV" in format_type:
                file_filter = "CSV File (*.csv)"
                default_ext = ".csv"
            else:
                file_filter = "Text File (*.txt)"
                default_ext = ".txt"

            # Show save file dialog
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Results",
                f"email_results{default_ext}",
                file_filter
            )

            if file_path:
                if not any(file_path.endswith(ext) for ext in ['.xlsx', '.csv', '.txt']):
                    file_path += default_ext

                if file_path.endswith('.xlsx'):
                    # Save as Excel
                    df.to_excel(file_path, index=False, sheet_name='Email List')
                elif file_path.endswith('.csv'):
                    # Save as CSV
                    df.to_csv(file_path, index=False, encoding='utf-8')
                else:
                    # Save as text file
                    with open(file_path, "w", encoding="utf-8") as file:
                        for email, source in self.current_results:
                            file.write(f"{email}, {source}\n")

                QMessageBox.information(self, "Success", f"Results successfully saved to {file_path}!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error occurred while saving file: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
