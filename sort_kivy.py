import os
import pandas as pd
import datetime
import re

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout  
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button

from kivy.uix.popup import Popup
from kivy.core.window import Window
from kivy.utils import get_color_from_hex
from plyer import filechooser  


# Set Kivy background color
Window.clearcolor = get_color_from_hex('#F0F4F8')

# Get desktop path and create SORT RESULT folder
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
sort_result_folder = os.path.join(desktop_path, "SORT RESULT")

if not os.path.exists(sort_result_folder):
    os.makedirs(sort_result_folder)


class ExcelFilterApp(App):
    def build(self):
        self.icon = r"C:\Goddie\Numbers-sort\icon.png"
        self.title = "Numbers Sorting Application"

        # Outer layout
        root = BoxLayout(orientation='vertical', padding=20, spacing=10)

        # --- Banner ---
        banner = Label(
            text="Filter and sort phone numbers based on names",
            size_hint_y=None,
            height=50,
            color=get_color_from_hex('#FFFFFF'),
            bold=True,
            font_size=20,
            padding=(10, 10),
        )
        with banner.canvas.before:
            from kivy.graphics import Color, Rectangle

            Color(*get_color_from_hex('#0056b3'))  # A darker shade of blue
            self.rect = Rectangle(size=banner.size, pos=banner.pos)
        banner.bind(size=self._update_rect, pos=self._update_rect)
        root.add_widget(banner)

        # --- Input Grid ---
        input_grid = GridLayout(cols=2, spacing=20, size_hint_y=0.5)

        # Folder Selection
        input_grid.add_widget(
            Label(text="Input path or browse folder:", color=get_color_from_hex('#333333'))
        )
        folder_input_layout = BoxLayout(orientation='horizontal', size_hint_x=1)
        self.folder_var = TextInput(multiline=False, size_hint_x=1.5)
        browse_button = Button(
            text="Browse Folder",
            size_hint_x=1,  # Adjusted size_hint_x
            background_color=get_color_from_hex('#007BFF'),
        )
        browse_button.bind(on_release=self.show_file_chooser)
        folder_input_layout.add_widget(self.folder_var)
        folder_input_layout.add_widget(browse_button)
        input_grid.add_widget(folder_input_layout)

        # Search Name Input
        input_grid.add_widget(
            Label(text="Input Name to search:", color=get_color_from_hex('#333333'))
        )
        self.search_name_var = TextInput(multiline=False)
        input_grid.add_widget(self.search_name_var)

        root.add_widget(input_grid)

       
        execute_button = Button(
            text="Execute Filter",
            size_hint_y=0.1,
            background_color=get_color_from_hex('#28A745'),
        )
        execute_button.bind(on_release=self.execute_filter)
        root.add_widget(execute_button)

        return root

    def _update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size

    def show_file_chooser(self, instance):
        try:
            os.environ['PLYER_DEFAULT_PROVIDER'] = 'win32'  
            filechooser.choose_dir(title="Pick a folder", on_selection=self.selected)  
        except Exception as e:
            print(f"Error in show_file_chooser: {e}")
            import traceback
            traceback.print_exc() 

    def selected(self, filename):
        try:
            if filename:
                self.update_folder_path(filename[0])
        except Exception as e:
            print(f"Error in selected: {e}")
            import traceback
            traceback.print_exc()  # Print the full traceback


    def update_folder_path(self, path):
        self.folder_var.text = path

    def execute_filter(self, instance):
        folder = self.folder_var.text
        search_name = self.search_name_var.text.strip().lower()
        output_file = os.path.join(
            sort_result_folder,
            f"{search_name}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )

        # DataFrame to store all filtered results for the search term
        filtered_data = pd.DataFrame(
            columns=['Credit Identity String', 'Customer Name']
        )

        # Define required columns
        required_columns = ['Credit Identity String', 'Customer Name']

        try:
            # Iterate over each file in the specified directory
            for filename in os.listdir(folder):
                filepath = os.path.join(folder, filename)

                # Handle Excel files
                if filename.endswith('.xlsx') or filename.endswith('.xlsm'):
                    xls = pd.ExcelFile(filepath)

                    for sheet_name in xls.sheet_names:
                        df = xls.parse(sheet_name)
                        print(f"Checking file: {filename}, sheet: {sheet_name}")

                        # Normalize column names by stripping whitespace
                        df.columns = df.columns.str.strip()

                        # Debugging: Print available columns
                        print(
                            f"Columns in {filename}, sheet {sheet_name}: {df.columns.tolist()}"
                        )

                        # Check for the required columns
                        actual_columns = df.columns.str.strip().str.lower()

                        if all(
                            col.lower() in actual_columns for col in required_columns
                        ):
                            # Filter rows based on the search_name in 'Customer Name' column
                            regex_pattern = rf'\b{re.escape(search_name)}\b'
                            filtered_rows = df[
                                df['Customer Name'].str.match(
                                    regex_pattern, case=False, na=False
                                )
                            ]

                            print(
                                f"Number of filtered rows for '{search_name}': {len(filtered_rows)}"
                            )

                            if not filtered_rows.empty:
                                available_columns = [
                                    col
                                    for col in ['Credit Identity String', 'Customer Name']
                                    if col in df.columns
                                ]
                                filtered_data = pd.concat(
                                    [filtered_data, filtered_rows[available_columns]],
                                    ignore_index=True,
                                )
                            else:
                                missing_cols = (
                                    set(required_columns) - set(actual_columns)
                                )
                                print(
                                    f"Required columns not found in {filename} - {sheet_name}: {missing_cols}"
                                )
                                print("First few rows of the DataFrame:")
                                print(df.head())

                # Handle CSV files
                elif filename.endswith('.csv'):
                    try:
                        df = pd.read_csv(
                            filepath, header=0, on_bad_lines='skip'
                        )
                        print(f"Checking file: {filename} (CSV)")

                        df.columns = df.columns.str.strip()

                        print(
                            f"Columns in {filename} (CSV): {df.columns.tolist()}"
                        )

                        actual_columns = df.columns.str.strip().str.lower()

                        if all(
                            col.lower() in actual_columns for col in required_columns
                        ):
                            regex_pattern = rf'\b{re.escape(search_name)}\b'
                            filtered_rows = df[
                                df['Customer Name'].str.match(
                                    regex_pattern, case=False, na=False
                                )
                            ]

                            print(
                                f"Number of filtered rows for '{search_name}': {len(filtered_rows)}"
                            )

                            if not filtered_rows.empty:
                                available_columns = [
                                    col
                                    for col in ['Credit Identity String', 'Customer Name']
                                    if col in df.columns
                                ]
                                filtered_data = pd.concat(
                                    [filtered_data, filtered_rows[available_columns]],
                                    ignore_index=True,
                                )
                            else:
                                missing_cols = (
                                    set(required_columns) - set(actual_columns)
                                )
                                print(
                                    f"Required columns not found in {filename} (CSV): {missing_cols}"
                                )
                                print("First few rows of the DataFrame:")
                                print(df.head())

                    except pd.errors.ParserError as e:
                        print(f"Error parsing {filename}: {e}")

            # Save the filtered data for the specific search term in SORT RESULT folder
            if not filtered_data.empty:
                filtered_data.to_excel(output_file, index=False)
                self.show_popup(
                    "Success",
                    f"Filtered results for '{search_name}' saved to '{output_file}'.",
                )
            else:
                self.show_popup(
                    "No Results", f"No results found for '{search_name}'."
                )

        except Exception as e:
            self.show_popup("Error", str(e))

    def show_popup(self, title, message):
        popup = Popup(
            title=title,
            content=Label(text=message),
            size_hint=(None, None),
            size=(400, 200),
        )
        popup.open()


if __name__ == '__main__':
    ExcelFilterApp().run()
