from E3_COM import Job, E3
from functions import *
from nom_of_element import *
from win32com import client
import os
import datetime
import sys


def make_note(note):
    split_note = note.split(' - ', 1)
    link = split_note[0]
    new_note = split_note[1] if len(split_note) == 2 else ''
    num = re.search(r'\d+', link).group()
    if new_note:
        new_note = num + ' ' + new_note
    return link, new_note


def create_end_notes(device, end_notes):
    note = device.get('note')
    inside_devs = device.get('inside_devs')
    if 'См. прим.' in note:
            link, new_note = make_note(note)
            if new_note and new_note not in end_notes:
                end_notes.append(new_note)
            device['note'] = link
    if inside_devs:
        for dev in inside_devs:
            create_end_notes(dev, end_notes)


def print_device(device, start_row, excel):    
    sheet = excel.ActiveSheet
    excel.Columns(2).ColumnWidth = 70
    excel.Columns(4).ColumnWidth = 15
    excel.Columns(5).ColumnWidth = 15
    current_row = start_row
    sheet.Rows(current_row).Cells(1).Value = device.get('ref')
    sheet.Rows(current_row).Cells(2).Value = device.get('name')
    sheet.Rows(current_row).Cells(3).Value = device.get('cnt')
    sheet.Rows(current_row).Cells(4).Value = device.get('note')
    sheet.Rows(current_row).Cells(5).Value = device.get('list_position')
    type_rec = device.get('type_rec')
    if type_rec:
        sheet.Rows(current_row).Cells(6).Value = type_rec
    current_row += 1
    inside_devs = device.get('inside_devs')
    if inside_devs:
        for inside_dev in inside_devs:
            if inside_dev.get('ref'):
                current_row = print_device(inside_dev, current_row, excel) #+ 1
            else:
                current_row = print_device(inside_dev, current_row, excel)
    return current_row


def print_devices(device_list, excel):
    selection = excel.Application.Selection
    sheet = excel.ActiveSheet
    start_row = selection.Cells(1).Row
    end_list = []
    end_notes = []
    for device in device_list:
        create_end_notes(device, end_notes)
        inside_devs = device.get('inside_devs')
        if inside_devs:
            for inside_dev in inside_devs:
                if inside_dev.get('ref'):
                    end_dev = device.copy()
                    end_dev['type_rec'] = 'X'
                    end_list.append(end_dev)
                    device.pop('inside_devs')
                    break
    
    for device in device_list:
        start_row = print_device(device, start_row, excel) + 1
    group_devices(end_list)
    for dev in end_list:
        start_row = print_device(dev, start_row, excel) + 1
    if end_notes:
        sheet.Rows(start_row).Cells(2).Value = 'Примечания:'
        sheet.Rows(start_row).Cells(5).Value = '1'
        for note in sorted(end_notes):
            start_row += 1
            sheet.Rows(start_row).Cells(2).Value = note
            

def create_excel_list():
    sheet = Job.create_sheet()
    doc = Job.create_doc()

    sheet.SetId(Job.get_active_sheet_id())

    # End script if there is no opened sheets
    if sheet.GetId() == 0:
        E3.put_warning(1, 'No opened sheets')
        del sheet
        del Job.job
        del E3.e3
        sys.exit()

    # Parametres of scheme from working sheet
    S_assign = sheet.GetAssignment() 
    S_loc = sheet.GetLocation()
    S_obz = sheet.GetAttributeValue ("DECIMALN")

    E3.put_info(0, f"________Start________\n{S_obz}    Loc: {S_loc}         Assign: {S_assign}")
    now = datetime.datetime.now()
    excel = client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False
    excel_workbook = excel.Workbooks.Add()
    excel_name = 'C:\\temp\\' +  S_obz.replace('Э','ПЭ') + '.xlsx'
    doc_title = S_obz.replace('Э','ПЭ') + ' ' + now.strftime("%Y-%m-%d %H:%M:%S")
    if os.path.exists(excel_name):
        os.remove(excel_name)
    
    device_list = nom_of_element()
    print_devices(device_list, excel)   
    excel.ActiveWorkbook.SaveAs(excel_name)
    excel.Quit()
    #Excel = None
    del excel
    # excel = None
    doc.Create(0, doc_title, excel_name)		
    doc.SetAssignment(S_assign) 
    doc.SetLocation(S_loc)
    doc.display()
    E3.put_info(0, "End script")
    del Job.job
    del E3.e3
    sys.exit()
