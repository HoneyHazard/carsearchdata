#!/bin/python3
import os
import subprocess
import csv
import openpyxl as px
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from tempfile import NamedTemporaryFile
from datetime import datetime

def open_url_in_chromium(url):
    subprocess.Popen(['chromium-browser', url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def get_input_from_user(prompt):
    return input(f"{prompt}: ")

def get_hours_input():
    with NamedTemporaryFile(suffix="_hours.txt") as temp_file:
        subprocess.call(['nano', temp_file.name])
        with open(temp_file.name, 'r') as f:
            hours = f.read().strip().replace('\n', ', ')
    return hours
    
def get_location_input():
    with NamedTemporaryFile(suffix="_address.txt") as temp_file:
        subprocess.call(['nano', temp_file.name])
        with open(temp_file.name, 'r') as f:
            location = f.read().strip().replace('\n', ', ')
    return location

def write_to_excel(data, filename):
    if os.path.exists(filename):
        wb = px.load_workbook(filename)
        ws = wb.active
    else:
        wb = px.Workbook()
        ws = wb.active
        ws.append(['Car', 'Mileage', 'Price', 'Link', 'Drive', 'Engine', 'VIN', 'MPG', 'Accidents/Damage/Title', 
                   'Condition', 'Company', 'Website', 'Location', 'Latitude', 'Longitude', 'Hours', 'Contact', 'Phone', 'Email', 'Status'])
        bold_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = bold_font
    
    ws.append(data)
    wb.save(filename)

def main():
    urls_file = input("Enter the path to the file containing URLs: ")

    with open(urls_file, 'r') as f:
        urls = f.readlines()

    for url in urls:
        url = url.strip()
        open_url_in_chromium(url)
        
        car_info = []
        car_info.append(get_input_from_user("year and model"))
        car_info.append(get_input_from_user("mileage"))
        car_info.append(get_input_from_user("price"))
        car_info.append(get_input_from_user("link"))
        car_info.append(get_input_from_user("drive"))
        car_info.append(get_input_from_user("engine"))
        car_info.append(get_input_from_user("VIN"))
        car_info.append(get_input_from_user("MPG"))
        car_info.append(get_input_from_user("accidents/damage/title"))
        car_info.append(get_input_from_user("condition"))
        car_info.append(get_input_from_user("company/person"))
        car_info.append(get_input_from_user("website"))
        car_info.append(get_location_input())
        car_info.append("latitude")
        car_info.append("longitude")
        car_info.append(get_hours_input())
        car_info.append(get_input_from_user("contact"))
        car_info.append(get_input_from_user("phone"))
        car_info.append(get_input_from_user("email"))
        car_info.append(get_input_from_user("status"))

        write_to_excel(car_info, 'output.xlsx')

if __name__ == "__main__":
    main()

