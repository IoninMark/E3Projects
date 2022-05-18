# -*- coding: utf8 -*-
from win32com import client


class Database():
    database = None
    def __init__(self):
        self.database = client.Dispatch("ADODB.Connection")
            
    def open(self, db):
        self.database.Open(db)

    def execute(self, str):
         return self.database.Execute(str)
    #def create_db_object():
        #return client.Dispatch("ADODB.Connection")



class E3():
    e3 = client.Dispatch("CT.Application")

    @classmethod
    def create_project(cls):
        return cls.e3.CreateJobObject()

    @classmethod
    def get_cmp_database(cls):
        return cls.e3.GetComponentDatabase()
    
    
    @classmethod
    def put_info(cls, pop_up: bool, text: str):
        cls.e3.PutInfo(pop_up, text)

    @classmethod
    def put_message(cls, text: str):
        cls.e3.PutMessage(text)


    @classmethod
    def put_warning(cls, pop_up: bool, text: str):
        cls.e3.PutWarning(pop_up, text)
    
    def __init__(self):
        pass

class BaseObject():

    def __init__(self,):
        self.object = None

    def get_id(self):
        return self.object.GetId()

    def set_id(self, object_id: int):
        self.object.SetId(object_id)

    def get_name(self):
        return self.object.GetName()

    def set_name(self, new_name):
        self.object.SetName(new_name)

    def get_attribute(self, attribute_name):
        return self.object.GetAttributeValue(attribute_name)

    def set_attribute(self, attribute_name, value):
        self.object.SetAttributeValue(attribute_name, value)
        
    def get_attribute_ids(self):
        _, result =  self.object.GetAttributeIds()
        return result[1:]

    id = property(get_id, set_id)
    name = property(get_name, set_name)



class Job():
    job = E3.create_project()

    def __init__(self):
        pass

    @classmethod
    def create_sheet(cls):
        return cls.job.CreateSheetObject()

    @classmethod
    def create_doc(cls):
        return cls.job.CreateExternalDocumentObject()

    @classmethod
    def create_text(cls):
        return cls.job.CreateTextObject()    

    @classmethod
    def create_device(cls):
        return cls.job.CreateDeviceObject()


    @classmethod
    def create_symbol(cls):
        return cls.job.CreateSymbolObject()


    @classmethod
    def create_pin(cls):
        return cls.job.CreatePinObject()
        
    @classmethod
    def create_component(cls):
        return cls.job.CreateComponentObject()

    @classmethod
    def create_attribute(cls):
        return cls.job.CreateAttributeObject()

    @classmethod
    def create_graph(cls):
        return cls.job.CreateGraphObject()

    @classmethod
    def create_field(cls):
        return cls.job.CreateFieldObject()


    @classmethod
    def get_sheet_id_selected_in_tree(cls):
        _, result = cls.job.GetTreeSelectedSheetIds()
        return result[1:]

    @classmethod
    def get_all_device_ids(cls):
        _, result = cls.job.GetAllDeviceIds()
        return result[1:]
   
    @classmethod
    def get_all_component_ids(cls):
        _, result = cls.job.GetAllComponentIds()
        return result[1:]
   
    @classmethod
    def get_sheet_ids(cls):
        _, result = cls.job.GetSheetIds()
        return result[1:]

    @classmethod
    def get_active_sheet_id(cls):
        return cls.job.GetActiveSheetId()
    
    @classmethod
    def get_sheet_id_selected_in_tree_by_folder(cls):
        _, result = cls.job.GetTreeSelectedSheetIdsByFolder()
        return result[1:]
    
    @classmethod
    def get_symbol_id_selected_in_tree(cls):
        _, result = cls.job.GetTreeSelectedSymbolIds()
        return result[1:]    


    @classmethod
    def get_symbol_id_selected_in_sheet(cls):
        _, result = cls.job.GetSelectedSymbolIds()
        return result[1:]

    @classmethod
    def get_gid_of_id(cls, id):
        return cls.job.GetGidOfId(id)



class Device():
    device = Job.create_device()

    @classmethod
    def get_symbol_ids_by_device_id(cls, device_id):
        cls.device.SetId(device_id)
        _, _result = cls.device.GetSymbolIds()
        return _result[1:]

    def __init__(self, device_id):
        self.device = Job.create_device()
        self.device.SetId(device_id)

    def set_id(self, device_id):
        self.device.SetId(device_id)

    def get_id(self):
        return self.device.GetId()

    def get_ref_des(self):
        return self.device.GetName()

    def get_assignment(self):
        return self.device.GetAssignment()

    def get_name(self):
        pass


    id = property(get_id, set_id)



    def get_symbol_ids(self):
        _, _result = self.device.GetSymbolIds()
        return _result[1:]

    def get_device_ids(self):
        _, _result = self.device.GetDeviceIds()
        return _result[1:]

    def get_symbols(self):
        _symbols_id = self.get_symbol_ids()
        _symbols = []
        for symbol_id in _symbols_id:
            _symbols.append(Symbol(symbol_id))
        return _symbols


    





class Sheet():
    pass

class Pin():
    pass

class Symbol(BaseObject):
    # symbol = Job.create_symbol()

    # def __init__(self, symbol_id):
    #     self.symbol = Job.create_symbol()
    #     self.id = symbol_id

    # def get_id(self):
    #     return self.symbol.GetId()

    # def set_id(self, symbol_id):
    #     self.symbol.SetId(symbol_id)

    # id = property(get_id, set_id)

    def __init__(self, object_id):
        self.object = Job.create_symbol()
        self.id = object_id

class Comparator:
    def __init__(self, ref_des = None, assignment = None, location = None):
        self.ref_des = ref_des
        self.assignment = assignment
        self.location = location
    
    def compare_ref_des(self, compare_object: Device):
        if compare_object.get_ref_des() == self.ref_des:
            return True
        else:
            return False

    def compare_assignment(self, compare_object: Device):
        if compare_object.get_assignment() == self.assignment:
            return True
        else:
            return False
    




