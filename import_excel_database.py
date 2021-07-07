from __future__ import print_function
from scriptengine import *


#from System.Runtime.InteropServices import Marshal
#Excel = Marshal.GetActiveObject("Excel.Application")
database = system.ui.open_file_dialog(title = "Choose Excel database to import", filter = "Excel files (*.xlsl)|*.xlsx")

print("Opening {}".format(database))

class Worksheet:
    def __init__(self, backend):
        self.backend = backend
        dig = self.backend
        ids = dig.Columns("D")
        print(ids.Cells(1,1).Value())
        print(ids.Cells(2,1).Value())
        print(ids.Cells(3,1).Value())
       

    def get_column_values(self, column):
        return self._filter_out_cells_with_no_values(self.backend.Columns(column).Cells())

    def get_column_values_limited(self, column, max_count):
        values = [c for c in self.backend.Columns(column).Cells()]
        if len(values) > max_count:
            return list(values[0:max_count])
        else:
            return values

    def count_column_values(self, column):
        dig = self.backend
        ids = dig.Columns("D")
        print(ids.Cells(1,1).Value())
        print(ids.Cells(2,1).Value())
        print(ids.Cells(3,1).Value())

        values = self._filter_out_cells_with_no_values(self.backend.Columns(column).Cells())
        return len(values)

    def get_row_values(self, row):
        dig = self.backend
        ids = dig.Columns("D")
        print(ids.Cells(1,1).Value())
        print(ids.Cells(2,1).Value())
        print(ids.Cells(3,1).Value())

        dig = self.backend
        ids = dig.Columns(row)
        print(ids.Cells(1,1).Value())
        print(ids.Cells(2,1).Value())
        print(ids.Cells(3,1).Value())
        return self._filter_out_cells_with_no_values(self.backend.Rows(row).Cells())
        
    def _filter_out_cells_with_no_values(self, cells):
        cells_with_values = []
        for cell in cells:
            if cell is None or cell == "":
                break
            cells_with_values.append(cell)
        
        return cells_with_values

    def get_name(self):
        return self.backend.Name

class Workbook:
    def __init__(self, backend):
        self.backend = backend

    def get_worksheet_by_name(self, name):
        try:
            return self._get_worksheet_by_name(name)
        except:
            raise Exception("Could not find worksheet: {}".format(name))

    def _get_worksheet_by_name(self, name):
        return Worksheet(self.backend.Sheets(name))

class Excel:
    def __init__(self):
        import clr
        clr.AddReference("Microsoft.Office.Interop.Excel")
        import Microsoft.Office.Interop.Excel as Excel    
        self.excel = Excel.ApplicationClass()
        self.excel.Visible = False
        self.workbooks = []

    def __del__(self):
        for wb in self.workbooks:
            wb.backend.Close()
        self.excel.Quit()

    def open(self, path):
        try:
           return self._open(path)
        except:
            raise Exception("Could not open workbook: {}".format(path))

    def _open(self, path):
        workbook = Workbook(self.excel.Workbooks.Open(path))
        self.workbooks.append(workbook)
        return workbook
        

def raise_by_number_of_decimals(value, decimals):
    return int(float(value)*pow(10, int(decimals)))

class DatabaseEntryData:
    def __init__(self, excluded, type, id, name):
        self.excluded = excluded
        self.type = type
        self.id = id
        self.name = name

class DigitalInputOutput(DatabaseEntryData):
    def __init__(self, init_database_method_string, excluded, type, id, name, channel, pin, location):
        DatabaseEntryData.__init__(self, excluded, type, id, name)
        self.init_database_method_string = init_database_method_string
        self.channel = channel
        self.pin = pin
        self.location = location
        self.comment = "{}, Channel: {}, Pin: {}, Location: {}".format(type, channel, pin, location)

    def get_init_database_string(self):
        return "{}(DatabaseID.{}, epirocTypes.Location.{}, {}, {});\n".format(
            self.init_database_method_string,
            DatabaseNameFormater.format(self.name), 
            self.location, 
            self.channel, 
            self.pin)

class DigitalOutput(DigitalInputOutput):
    def __init__(self, excluded, id, name, channel, pin, location):
        DatabaseEntryData.__init__(self, "init_digital_output", excluded, "PhysicalDigitalOutput", id, name, channel, pin, location)
    
class Parameter(DatabaseEntryData):
    def __init__(self, id, name, value, max_value, min_value, num_dec, unit, comment):
        DatabaseEntryData.__init__(self, False, "Parameter", id, name)
        self.value = raise_by_number_of_decimals(value, num_dec)
        self.max_value = raise_by_number_of_decimals(max_value, num_dec)
        self.min_value = raise_by_number_of_decimals(min_value, num_dec)
        self.num_dec = int(num_dec)
        self.unit = unit
        self.comment = "{}, Default: {}, Max: {}, Min: {}, Num Decimals: {}, Unit: {}".format(comment, value, max_value, min_value, int(num_dec), unit)

    def get_init_database_string(self):
        return "init_parameter(DatabaseID.{}, {}, {}, {}, {}, epirocTypes.Units.{});\n".format(
            DatabaseNameFormater.format(self.name), 
            self.value, 
            self.max_value, 
            self.min_value, 
            self.num_dec, 
            self.unit)


class DatabaseEntryExtractorBase:
    def __init__(self, worksheet):
        self.worksheet = worksheet
        print("Get", worksheet.get_name())
        self.top_row = worksheet.get_row_values(1)
        print(self.top_row)
        self.number_of_entries_on_sheet = self.worksheet.count_column_values(self._find_data_name_index("name") + 1)
   
    def _get_data_array(self, data_name):
        data_index = self._find_data_name_index(data_name)
        return self._remove_top_row(self.worksheet.get_column_values_limited(data_index+1, self.number_of_entries_on_sheet))
        
    def _find_data_name_index(self, data_name):
        try:
            return self.top_row.index(data_name)
        except:
            raise Exception("Could not find {} in {}".format(data_name, self.worksheet.get_name()))

    def _remove_top_row(self, column):
        del column[0]
        return column


class ParameterExtractor(DatabaseEntryExtractorBase): 
    def __init__(self, worksheet):
        DatabaseEntryExtractorBase.__init__(self, worksheet)

    def extract_all_parameters(self):
        data_arrays = self._get_all_parameter_data_arrays()
        return self._create_parameters_from_data_arrays(data_arrays)
        
    def _get_all_parameter_data_arrays(self):
        class DataArrays(object): pass
        data_arrays = DataArrays()
        data_arrays.names = self._get_data_array("name")
        data_arrays.values = self._get_data_array("value")
        data_arrays.max_values = self._get_data_array("max_value")
        data_arrays.min_values= self._get_data_array("min_value")
        data_arrays.num_decs = self._get_data_array("num_dec")
        data_arrays.units = self._get_data_array("unit")
        data_arrays.comments = self._get_data_array("Comment")   
        return data_arrays
    
    def _create_parameters_from_data_arrays(self, data_arrays):
        parameters = []
        for id, name in enumerate(data_arrays.names):
            parameter = Parameter(id, name, data_arrays.values[id], data_arrays.max_values[id], data_arrays.min_values[id], data_arrays.num_decs[id], data_arrays.units[id], data_arrays.comments[id])
            parameters.append(parameter)

        return parameters


class DigitalOutputExtractor(DatabaseEntryExtractorBase):      
    def __init__(self, worksheet):
        DatabaseEntryExtractorBase.__init__(self, worksheet)
          
    def extract_all_digial_outputs(self):
        data_arrays = self._get_all_digial_output_data_arrays()
        return self._create_digial_output_from_data_arrays(data_arrays)
 
    def _get_all_digial_output_data_arrays(self):
        class DataArrays(object): pass
        data_arrays = DataArrays()
        data_arrays.excluded = self._get_data_array("Excluded:")
        data_arrays.names = self._get_data_array("name")
        data_arrays.channel = self._get_data_array("channel")
        data_arrays.pin = self._get_data_array("pin")
        data_arrays.location = self._get_data_array("location")
        return data_arrays

    def _create_digial_outputs_from_data_arrays(self, data_arrays):
        entries = []
        for id, name in enumerate(data_arrays.names):
            entry = DigitalOutput(data_arrays.excluded[id] != "", id, name, data_arrays.channel[id], data_arrays.pin[id], data_arrays.location[id])
            entries.append(entry)
       
        return entries

class DatabaseExtractor:
    @staticmethod
    def extract(workbook):
        class Entries(object): pass
        entries = Entries()
        entries.parameters = ParameterExtractor(workbook.get_worksheet_by_name("Parameters")).extract_all_parameters()
        entries.digital_outputs = DigitalOutputExtractor(workbook.get_worksheet_by_name("OutputsDig")).extract_all_digial_outputs()
        return entries

class DatabaseNameFormater:
    @staticmethod
    def format(name):
        return name.replace(" ", "_")

class DatabaseIDFormater:
    @staticmethod
    def format(to_format):
        return DatabaseIDFormater._format(to_format)

    @staticmethod
    def _format(to_format):
        return "\n".join(
            "{} := SHL(TO_DWORD(DatabaseType.{}), 16) OR {}{}\t//{}".format(
                DatabaseNameFormater.format(s.name),
                s.type,
                s.id,
                "" if i+1 == len(to_format) else ",",
                s.comment
            ) for i, s in enumerate(to_format))


class InitDatabaseFormater:
    @staticmethod
    def format(to_format):
        return "".join(
            s.get_init_database_string()
            for i, s in enumerate(to_format))

class ParametersPOUFormater:
    @staticmethod
    def format_declaration(to_format):
        return "\n".join(
            "{}_{}_{} : {}; \t//{}".format(
                DatabaseNameFormater.format(s.name),
                s.unit,
                s.num_dec,
                "BOOL" if s.max_value == 1 else "UINT",
                s.comment
            ) for i, s in enumerate(to_format))

    @staticmethod
    def format_definition(to_format):
        return "\n".join(
            "\t{}_{}_{} := ParameterGet{}(DatabaseID.{});".format(
                DatabaseNameFormater.format(s.name),
                s.unit,
                s.num_dec,
                "BOOL" if s.max_value == 1 else "UINT",
                DatabaseNameFormater.format(s.name)
            ) for i, s in enumerate(to_format))

try:
    excel = Excel()
    wb = excel.open(database)
    ws = wb.get_worksheet_by_name("Parameters")
    entries = DatabaseExtractor.extract(wb)
    print(DatabaseIDFormater.format(entries.parameters))
    print(DatabaseIDFormater.format(entries.digital_outputs))
    print(InitDatabaseFormater.format(entries.parameters))
    print(ParametersPOUFormater.format_declaration(entries.parameters))
    print(ParametersPOUFormater.format_definition(entries.parameters))
except:
    raise