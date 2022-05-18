from E3_COM import Job, E3
from win32com import client
from natsort import natsorted
import re


# Function returns the number in the end of alphanumeric string 'a1' -> 1 
def get_end_number_from_str(ref_str):
    index = len(ref_str) - 1
    while ref_str[index].isdigit():
        if index:
            index = index - 1
        else:
            break
    return int(ref_str[index + 1:]) if index + 1 < len(ref_str) else 0


# Function returns grouped reference of devices
def group_ref(ref, new_ref):
    split_comma = ref.split(', ')
    split_dash = ref.split(' - ')
    split_dash_new = new_ref.split(' - ')
    if len(split_comma) == 1:
        if len(split_dash_new) == 1:
            if len(split_dash) == 1:
                return ref + ', ' + new_ref if ref else new_ref
            if get_end_number_from_str(ref) == get_end_number_from_str(new_ref) - 1:
                return split_dash[0] + ' - ' + new_ref
            return ref + ', ' + new_ref if ref else new_ref
        else:
            if get_end_number_from_str(ref) == get_end_number_from_str(split_dash_new[0]) - 1:
                return ref.replace(split_dash[-1], split_dash_new[-1])
            else:
                return ref + ', ' + new_ref
    else:
        if len(split_dash_new) == 1:
            if get_end_number_from_str(split_comma[-2]) + 1 == get_end_number_from_str(split_comma[-1]) == get_end_number_from_str(new_ref) - 1:
                return ref.replace(', ' + split_comma[-1], '') + ' - ' + new_ref
            else:
                return ref + ', ' + new_ref if ref else new_ref
        else:
            if get_end_number_from_str(ref) == get_end_number_from_str(split_dash_new[0]) - 1:
                return ref.replace(split_dash[-1], split_dash_new[-1])
            else:
                return ref + ', ' + new_ref


# Function sorts the list of elements by name and ref
def sort_list(items):
    sorted_items = sorted(items, key = lambda k: k['name'])
    sorted_items = natsorted(sorted_items, key = lambda k: k['ref'])
    for item in sorted_items:
        inside_items = item.get('inside_devs')
        if inside_items:
            item['inside_devs'] = sort_list(inside_items)
    return sorted_items


# Function groups devices by their structure
def group_devices(dev_list):
    i = 0
    while i < len(dev_list) - 1:
        d1_dict = {
            'name': dev_list[i].get('name'),
            'devs': dev_list[i].get('inside_devs'),
            'note': dev_list[i].get('note') if dev_list[i].get('name').find('Клемма') == -1 else ''
            }
        d2_dict = {
            'name': dev_list[i+1].get('name'),
            'devs': dev_list[i+1].get('inside_devs'),
            'note': dev_list[i+1].get('note') if dev_list[i+1].get('name').find('Клемма') == -1 else ''
            }
        if d1_dict == d2_dict:
            dev_list[i]['cnt'] += 1
            if dev_list[i+1]['list_position']:
                dev_list[i]['list_position'] = dev_list[i+1]['list_position']
            if dev_list[i]['note'] != dev_list[i+1]['note']:
                for note in dev_list[i+1]['note'].split(', '):
                    dev_list[i]['note'] = group_ref(dev_list[i]['note'], note)
            if dev_list[i]['ref'] and dev_list[i+1]['ref']:
                dev_list[i]['ref'] = group_ref(dev_list[i]['ref'], dev_list[i+1]['ref']) 
            dev_list.pop(i+1)
            i = i - 1
        i += 1
    for dev in dev_list:
        inside_devs = dev.get('inside_devs')
        if inside_devs:
            group_devices(inside_devs)


# Function makes request to the sql database for add_part.
# Returs False and Error if add_part doesn't exist in database
# Returns False if add part has atrribute not to show in list
# Returns True and Name if ok
def get_part_from_database(component_name):
    database = E3.get_cmp_database()
    db_connection = client.Dispatch("ADODB.Connection")
    db_connection.Open(database)
    request1 = f"SELECT AttributeValue FROM ComponentAttribute WHERE ENTRY= '{component_name}' AND AttributeName= 'imbase_name' "
    request2 = f"SELECT AttributeValue FROM ComponentAttribute WHERE ENTRY= '{component_name}' AND AttributeName= 'DO_NOT_ADD_TO_LIST' "
    request3 = f"SELECT AttributeValue FROM ComponentAttribute WHERE ENTRY= '{component_name}' AND AttributeName= 'code_max' "
    rs1 = db_connection.execute(request1)
    rs2 = db_connection.execute(request2)
    rs3 = db_connection.execute(request3)
    if rs1.EOF:
        return False, "Error"
    if not rs2.EOF:
        rs2.MoveFirst()
        if rs2.Fields('AttributeValue').Value:
            return False, ""
    if not rs3.EOF:
        rs3.MoveFirst()   
    rs1.MoveFirst()
    result = True, rs1.Fields('AttributeValue').Value, rs3.Fields('AttributeValue').Value
    db_connection.Close()
    return result


# Function to get additional devices
# returns  list[] of additional devices 
# every device in result list is dictionary with same structure as input device
def get_add_parts(dev):
    result = []
    device = {}
    att = Job.create_attribute()    
    _, att_ids = dev.GetAttributeIds()
    att_ids = att_ids[1:]
    for att_id in att_ids:
        att.SetId(att_id)
        if att.GetName() == 'Дополнительная часть':
            cmp_prmtrs =  att.GetValue().rsplit(':')[1:]
            cmp_name = cmp_prmtrs[1]
            cmp_cnt = cmp_prmtrs[0]
            if len(cmp_prmtrs) > 2:
                cmp_note = ''.join(cmp_prmtrs[2:])
            else:
                cmp_note = ''
            part_exist, name, code = get_part_from_database(cmp_name)
            if part_exist:
                device = {
                    'ref': dev.GetName()[1:],
                    'name': name,
                    'cnt': cmp_cnt,
                    'note': cmp_note,
                    'inside_devs': [],
                    'list_position': code
                    } 
                result.append(device)
            elif name == 'Error':
                E3.put_info(0, f"Add_part doesn't exist in database!!! {dev.Getname()}   part: {cmp_name}")
    return result


# Function, that returns the note of device
def get_dev_note(dev):
    res_note = ''
    notes_to_sort = []
    pin = Job.create_pin()
    if dev.IsTerminal() and dev.GetName().find('RU') == -1: 
        _, pin_ids = dev.GetPinIds()
        pin_ids = pin_ids[1:]
        for pin_id in pin_ids:
            pin.SetId(pin_id)
            note = f':{pin.GetName()}'
            if note not in notes_to_sort:
                notes_to_sort.append(note)
        for item in natsorted(notes_to_sort):    
            res_note = group_ref(res_note, item)
    else:
        note = dev.GetAttributeValue("primechanie_PE")
        if note:
            res_note += note 
    return res_note


# Function to open up the device
# returns yhe list[] of devices which input device contains
# every device in result list is dictionary with same structure as input device
def get_inside_devs(dev_id):
    result = []
    dev = Job.create_device()
    cmp = Job.create_component()
    dev.SetId(dev_id)
    root_name = dev.GetName()
    # Open the inside device if it is Additional Part
    if dev.GetAttributeValue("AdditionalPart"):     
        result.extend(get_add_parts(dev))

    # Open the inside device if it is Assembly or Terminal Block
    if dev.IsTerminalBlock() or dev.IsAssembly() and dev.GetComponentName() == '':
        _, dev_ids = dev.GetDeviceIds()
        dev_ids = dev_ids[1:]
        for id in dev_ids:
            dev.SetId(id)
            cmp.SetId(id)
            if dev.GetAttributeValue("DO_NOT_ADD_TO_LIST") != '1':
                inside_devs = get_inside_devs(id)
                dev_name = dev.GetName()[1:]
                device = {
                    'ref': dev_name, #[dev_name.find('.') + 1:],
                    'name': cmp.GetAttributeValue("imbase_name"),
                    'cnt': 1,
                    'note': get_dev_note(dev),
                    'inside_devs': inside_devs,
                    'list_position': cmp.GetAttributeValue("code_max")
                }
                # If device isn't assembly but has add_part:
                # Move this device inside "virtual assembly" with same ref and name as device class
                #if cmp.GetAttributeValue("imbase_name") and inside_devs:
                    #new_inside_dev = device.copy()
                    #new_inside_dev['ref'] = ''
                    #new_inside_dev['list_position'] = ''
                    #device['name'] = cmp.GetAttributeValue("Class")
                    #device['note'] = ''
                    #device['inside_devs'] = [new_inside_dev]
                # If device is terminal block or assembly inside assebly without unique ref
                # Don't show assembly/terminal block, only take its inside devices
                if dev.IsTerminalBlock() and device['ref'] == '' and inside_devs:
                    result.extend(inside_devs)
                    continue

                result.append(device)  
    return result




# Function find name of the device
#def get_device_name(dev, cmp):
    names = {
        'XT': 'Блок клеммный',
        'XP': 'Вилка',
        'XS': 'Розетка',
        'SA': 'Переключатель',
        'SB': 'Кнопка',
        'RU': 'Блок варисторный',
        'FU': 'Предохранитель',
        'QF': 'Выключатель автоматический',
        'KM': 'Контактор'
    }
    #ref = re.sub('[\d|*]', '', dev.GetName()[1:])
    name = cmp.GetAttributeValue("imbase_name")
    #if dev.IsAssembly() or dev.IsTerminalBlock():
        #if not name:
            #name = dev.GetAttributeValue("imbase_name")
            #if not name:               
                #if ref in names: 
                    #name = names[ref]
    return name