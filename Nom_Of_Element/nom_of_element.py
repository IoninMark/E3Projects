from E3_COM import Job, E3
from functions import *
from win32com import client


def print_devs(devs):    
    for dev in devs:
        print('{}    {}    {}    {}    {}'.format(dev.get('ref'), dev.get('name'), dev.get('cnt'), dev.get('note'), dev.get('list_position')))
        if dev.get('inside_devs'):
            print('-------')
            print_devs(dev.get('inside_devs'))
            print('---------------------------------------------------------------------------------')

def nom_of_element(sheet_id=0):
    sheet = Job.create_sheet()
    dev = Job.create_device()
    cmp = Job.create_component()

    if not sheet_id:
        sheet.SetId(Job.get_active_sheet_id())
    else:
        sheet.SetId(sheet_id)

    # Parametres of scheme from working sheet
    S_assign = sheet.GetAssignment() 
    S_loc = sheet.GetLocation()
    #S_obz = sheet.GetAttributeValue ("DECIMALN")

    # Main part
    # Making list with dictionaries for each device
    # Dictionary structure:
    # dev = {'ref': dev_ref, 'name': dev_name, 'cnt': dev_cnt, 'note': dev_note, 'inside_devs': [dev1, dev2, dev3...], 'list_position': (-1, num)}
    # dev1, dev2 ... - have the same dictionary structure as dev
    # 'list_position': (-1, num)
    # -1 - new list
    # num - number of lines between devices
    devices = [] # result list with device's dictionaries
    dev_ids = Job.get_all_device_ids()
    for dev_id in dev_ids:
        dev.SetId(dev_id)
        cmp.SetId(dev_id)
        # Filter for devices in project
        dev_filter = [
            dev.GetLocation() == S_loc,
            dev.GetAssignment() == S_assign,
            dev.IsAssemblyPart() == 0,
            dev.IsCable() == 0,
            dev.IsWiregroup() == 0,
            dev.GetComponentName() != "Соединитель скрытый",
            dev.GetComponentName() != "Точка заземления",
            #dev.GetComponentName() != "Резерв",
            dev.IsTerminal() == 0 or dev.IsTerminalBlock() == 1,
            #dev.GetAttributeValue("imbase_name").find('Кабель') == -1,
            cmp.GetAttributeValue("imbase_name").find('Кабель') == -1,
            #dev.GetAttributeValue("imbase_name").find('Шлейф') == -1,
            cmp.GetAttributeValue("imbase_name").find('Шлейф') == -1,
            #dev.GetAttributeValue("imbase_name").find('Пигтейл') == -1,
            cmp.GetAttributeValue("imbase_name").find('Пигтейл') == -1,
            dev.GetName().find('WB') == -1,
            dev.GetName().find('ГП') == -1,
            dev.GetName().find('Внешние_перемычки') == -1,
            dev.GetAttributeValue("DO_NOT_ADD_TO_LIST") != '1',
            ]
        if all(dev_filter):
            #print(f'Device: {dev.GetName()}')
            inside_devs = get_inside_devs(dev.GetId())
            cmp.SetId(dev.GetId())
            cmp_name = cmp.GetAttributeValue("imbase_name")
            name = get_device_name(dev, cmp)
            list_position = dev.GetAttributeValue("poziciya_PE")
            device = {
                'ref': dev.GetName()[1:],
                'name': name,
                'cnt': 1,
                'note': get_dev_note(dev),
                'inside_devs': inside_devs,
                'list_position': list_position if list_position else '1'
            }
            # If device isn't assembly but has add_part:
            # Move this device inside "virtual assembly" with same ref and name as device class
            if cmp_name and inside_devs:
                new_inside_dev = device.copy()
                new_inside_dev['ref'] = ''
                new_inside_dev['list_position'] = ''
                device['name'] = cmp.GetAttributeValue("Class")
                device['note'] = ''
                device['inside_devs'] = [new_inside_dev]
            # If assembly in project has only 1 device in it without any devices inside it
            # don't show this device like assembly
            if len(inside_devs) == 1 and dev.IsAssembly() and cmp_name == '':
                inside_dev = device.get('inside_devs')[0]
                if inside_dev.get('ref') =='' and inside_dev.get('inside_devs') == []:
                    device['name'] = inside_dev.get('name')
                    device['note'] += inside_dev.get('note')
                    device['list_position'] += inside_dev.get('list_position')
                    device['inside_devs'] = []
                    E3.put_warning(0, f"Assembly of 1 item!!!  {device.get('ref')}")
            devices.append(device)
    
    devices.extend(get_fields_and_other_devs(sheet_id))
    sorted_devices = sort_list(devices)
    group_devices(sorted_devices)
    #print_devs(sorted_devices)
    return sorted_devices


# Function returns the name and id of first sheet wich matches the field or assignment\location
def get_first_sheet_name_id(field=0, assignment=0, location=0):
    sheet = Job.create_sheet()
    sheet_ids = Job.get_sheet_ids()
    if field:
        assign = field.GetDeviceAssignment()
        loc = field.GetDeviceLocation()
    else:
        assign = assignment
        loc = location
    for sheet_id in sheet_ids:
        sheet.SetId(sheet_id)
        sheet_filter = [
            sheet.GetAssignment() == assign,
            sheet.GetLocation() == loc,
            sheet.GetName() == '1',
            sheet.GetFormat().find('Схема') == 0
            ]
        if all(sheet_filter):
            break

    name = sheet.GetAttributeValue('NAIMENOV_LIST') + ' ' + sheet.GetAttributeValue('DECIMALN').split(' ')[0]
    return name, sheet.GetId()


def get_fields_and_other_devs(sheet_id=0):
    sheet = Job.create_sheet()
    dev = Job.create_device()
    field = Job.create_field()
    
    if not sheet_id:
        sheet.SetId(Job.get_active_sheet_id())
    else:
        sheet.SetId(sheet_id)

    # Parametres of scheme from working sheet
    S_assign = sheet.GetAssignment() 
    S_loc = sheet.GetLocation()
    S_obz = sheet.GetAttributeValue("DECIMALN")

    result = []
    fields = []
    fields_cache = []
    out_devs = []

    sheet_ids = Job.get_sheet_ids()
    for sheet_id in sheet_ids:
        sheet.SetId(sheet_id)
        if sheet.GetAssignment() == S_assign and sheet.GetLocation() == S_loc and sheet.GetAttributeValue("DECIMALN") == S_obz:
            _, graph_ids = sheet.GetGraphIds()
            for graph_id in graph_ids[1:]:
                field.SetId(graph_id)
                if field.GetId():
                    assign, loc = field.GetDeviceAssignment(), field.GetDeviceLocation()
                    if (assign, loc) not in fields_cache:
                        fields_cache.append((assign, loc))
                        fields.append(field.GetId())
                    else:
                        if field.GetAttributeValue("show_devices_in_PE") == '1' or field.GetAttributeValue("DO_NOT_ADD_TO_LIST") == '1':
                            fields.pop(fields_cache.index((assign, loc)))
                            fields.append(field.GetId())

            # Take all devices from scheme
            _, sym_ids = sheet.GetSymbolIds()
            for sym_id in sym_ids[1:]:
                dev.SetId(sym_id)
                id = dev.GetId()
                if id:
                    if dev.GetAssignment() != S_assign or dev.GetLocation() != S_loc:
                        if id not in out_devs:
                            out_devs.append(id)

    for field_id in fields:
        field.SetId(field_id)
        if field.GetAttributeValue("DO_NOT_ADD_TO_LIST") != '1':
            i = 0
            while i < len(out_devs):
                dev.SetId(out_devs[i])
                if dev.GetAssignment() == field.GetDeviceAssignment() and dev.GetLocation() == field.GetDeviceLocation():
                    out_devs.pop(i)
                else:
                    i += 1
            device_name, sht_id = get_first_sheet_name_id(field=field)
            list_position = field.GetAttributeValue("poziciya_PE")
            device = {
                'ref': field.GetDeviceName()[1:],
                'name': device_name,
                'cnt': 1,
                'note': field.GetAttributeValue("primechanie_PE"),
                'list_position': list_position if list_position else '1'
            }
            if field.GetAttributeValue("show_devices_in_PE") == '1':
                device['inside_devs'] = nom_of_element(sht_id)
            else:
                device['inside_devs'] = []
        else:
            continue

        result.append(device)
        
    if out_devs:
        E3.put_warning(0, 'There are devices from other schemes without fields')
    return result