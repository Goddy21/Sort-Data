from kivy.config import Config
Config.set('graphics', 'resizable', False)
Config.set('graphics', 'width', '600')
Config.set('graphics', 'height', '350')
Config.set('graphics', 'position', 'custom')
Config.set('graphics', 'left', '100')
Config.set('graphics', 'top', '100')

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
from kivy.uix.progressbar import ProgressBar
from kivy.clock import Clock
from plyer import filechooser  

Window.clearcolor = get_color_from_hex('#F0F4F8')
Window.size = (600, 350) 
Window.resizable = False

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
sort_result_folder = os.path.join(desktop_path, "SORT RESULT")
if not os.path.exists(sort_result_folder):
    os.makedirs(sort_result_folder)


class ExcelFilterApp(App):
    def build(self):
        self.icon = r"C:\Goddie\Numbers-sort\icon.png"
        self.title = "Mkononi Numbers-Sorting System"

        self.root_layout = BoxLayout(orientation='vertical', padding=20, spacing=10)

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
            Color(*get_color_from_hex('#0056b3'))
            self.rect = Rectangle(size=banner.size, pos=banner.pos)
        banner.bind(size=self._update_rect, pos=self._update_rect)
        self.root_layout.add_widget(banner)

        input_grid = GridLayout(cols=2, spacing=20, size_hint_y=0.5)

        input_grid.add_widget(
            Label(text="[b]Input path or browse folder[/b]:", color=get_color_from_hex('#333333'), font_size='15sp', markup=True)
        )
        folder_input_layout = BoxLayout(orientation='horizontal', size_hint_x=1)
        self.folder_var = TextInput(multiline=False, size_hint_x=1.5)
        browse_button = Button(text="Browse Folder", size_hint_x=1, background_color=get_color_from_hex('#007BFF'))
        browse_button.bind(on_release=self.show_file_chooser)
        folder_input_layout.add_widget(self.folder_var)
        folder_input_layout.add_widget(browse_button)
        input_grid.add_widget(folder_input_layout)

        input_grid.add_widget(
            Label(text="[b]Input Name to search(s):[/b]", color=get_color_from_hex('#333333'), font_size='15sp', markup=True)
        )
        self.search_name_var = TextInput(multiline=False)
        input_grid.add_widget(self.search_name_var)

        self.root_layout.add_widget(input_grid)

        self.execute_button = Button(text="Execute Filter", size_hint_y=0.1, background_color=get_color_from_hex('#28A745'))
        self.execute_button.bind(on_release=self.start_filtering)
        self.progress_bar = ProgressBar(max=100, value=0, size_hint_y=0.05)
        self.root_layout.add_widget(self.progress_bar)
        self.root_layout.add_widget(self.execute_button)

        return self.root_layout

    def _update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size

    def show_file_chooser(self, instance):
        try:
            os.environ['PLYER_DEFAULT_PROVIDER'] = 'win32'
            filechooser.choose_dir(title="Pick a folder", on_selection=self.selected)
        except Exception as e:
            print(f"Error in show_file_chooser: {e}")

    def selected(self, filename):
        if filename:
            self.folder_var.text = filename[0]

    def start_filtering(self, instance):
        self.progress_bar.value = 0
        self.execute_button.disabled = True

        folder = self.folder_var.text
        search_input = self.search_name_var.text.strip().lower()

        self.search_names = list(set(name.strip() for name in re.split(r'[,\n\r]+', search_input) if name.strip()))
        if not self.search_names:
            self.show_popup("Input Error", "Please enter at least one name to search.")
            self.execute_button.disabled = False
            return

        self.files_created = 0
        self.required_columns = ['Credit Identity String', 'Customer Name']
        self.total = len(self.search_names)
        self.index = 0
        self.folder = folder
        Clock.schedule_once(self.process_next_name, 0.1)

    def process_next_name(self, dt):
        if self.index >= self.total:
            self.execute_button.disabled = False
            if self.files_created > 0:
                self.show_popup("Success", f"Created {self.files_created} file(s) for matching name(s).")
            else:
                self.show_popup("No Matches", "No matches found for the provided name(s).")
            return

        search_name = self.search_names[self.index]
        filtered_data = pd.DataFrame(columns=self.required_columns)

        for filename in os.listdir(self.folder):
            filepath = os.path.join(self.folder, filename)

            try:
                if filename.endswith(('.xlsx', '.xlsm')):
                    xls = pd.ExcelFile(filepath)
                    for sheet_name in xls.sheet_names:
                        df = xls.parse(sheet_name)
                        df.columns = df.columns.str.strip()
                        if all(col.lower() in df.columns.str.lower() for col in self.required_columns):
                            pattern = rf'\b{re.escape(search_name)}\b'
                            rows = df[df['Customer Name'].str.contains(pattern, case=False, na=False, regex=True)]
                            if not rows.empty:
                                filtered_data = pd.concat(
                                    [filtered_data, rows[self.required_columns]], ignore_index=True)

                elif filename.endswith('.csv'):
                    df = pd.read_csv(filepath, on_bad_lines='skip')
                    df.columns = df.columns.str.strip()
                    if all(col.lower() in df.columns.str.lower() for col in self.required_columns):
                        pattern = rf'\b{re.escape(search_name)}\b'
                        rows = df[df['Customer Name'].str.contains(pattern, case=False, na=False, regex=True)]
                        if not rows.empty:
                            filtered_data = pd.concat(
                                [filtered_data, rows[self.required_columns]], ignore_index=True)

            except Exception as e:
                print(f"Error processing file {filename}: {e}")

        if not filtered_data.empty:
            output_file = os.path.join(sort_result_folder, f"{search_name}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            filtered_data.to_excel(output_file, index=False)
            self.files_created += 1
            print(f"Saved: {output_file}")

        self.index += 1
        self.progress_bar.value = (self.index / self.total) * 100
        Clock.schedule_once(self.process_next_name, 0.1)

    def show_popup(self, title, message):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        message_label = Label(text=message, text_size=(380, None), halign='center', valign='middle', size_hint=(1, 1))
        close_button = Button(text='Close', size_hint=(1, 0.3), font_size='18sp', bold=True, background_normal='',
                              background_color=(0.2, 0.6, 0.86, 1), color=(1, 1, 1, 1))
        popup = Popup(title=title, content=layout, size_hint=(None, None), size=(400, 200), auto_dismiss=False)
        close_button.bind(on_release=popup.dismiss)
        layout.add_widget(message_label)
        layout.add_widget(close_button)
        popup.open()


if __name__ == '__main__':
    ExcelFilterApp().run()
