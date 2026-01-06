from kivymd.app import MDApp
from kivy.uix.boxlayout import BoxLayout
from kivymd.uix.pickers import MDDatePicker
from datetime import date, datetime
import openpyxl
import os


class TransportLayout(BoxLayout):
    selected_date = ""

    def open_date_picker(self):
        picker = MDDatePicker(max_date=date.today())
        picker.bind(on_save=self.on_date_selected)
        picker.open()

    def on_date_selected(self, instance, selected_date, date_range):
        self.selected_date = selected_date.strftime("%d-%m-%Y")
        self.ids.date_lbl.text = "Date: " + self.selected_date

    def save_data(self):
        if not self.selected_date:
            self.selected_date = datetime.now().strftime("%d-%m-%Y")

        file_name = "transport_data.xlsx"

        if not os.path.exists(file_name):
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append([
                "Date",
                "Vehicle Number",
                "Driver Name",
                "Material Name",
                "Quantity",
                "Meter Reading",
                "Diesel",
                "Other Expense"
            ])
            wb.save(file_name)

        wb = openpyxl.load_workbook(file_name)
        sheet = wb.active

        sheet.append([
            self.selected_date,
            self.ids.vehicle.text,
            self.ids.driver.text,
            self.ids.material.text,
            self.ids.qty.text,
            self.ids.meter.text,
            self.ids.diesel.text,
            self.ids.other.text
        ])

        wb.save(file_name)
        self.clear_fields()

    def clear_fields(self):
        self.ids.vehicle.text = ""
        self.ids.driver.text = ""
        self.ids.material.text = ""
        self.ids.qty.text = ""
        self.ids.meter.text = ""
        self.ids.diesel.text = ""
        self.ids.other.text = ""


class TransportApp(MDApp):
    def build(self):
        self.title = "Transport"
        return TransportLayout()


if __name__ == "__main__":
    TransportApp().run()

