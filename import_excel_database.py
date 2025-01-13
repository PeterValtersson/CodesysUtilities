#!/usr/local/bin/python
# -*- coding: utf-8 -*-
from __future__ import print_function
from scriptengine import *


#from System.Runtime.InteropServices import Marshal
#Excel = Marshal.GetActiveObject("Excel.Application")


database = system.ui.open_file_dialog(title = "Choose Excel database to import", filter = "Excel files (*.xlsl)|*.xlsx", directory = "C:\Users\tmgpeva\OneDrive - Epiroc\Dokument\non vault repos\CodesysUtilities")
import_type = system.ui.choose(message = "What to import", options = ["Import to database library", "Import to database library - no textlists","Import FlexiROC interim database lib", "Import inputs and outputs for project"])[0]
IMPORT_LIB = 0
IMPORT_LIB_NO_TEXTLIST = 1
IMPORT_LIB_FLEXI_INTERIM = 2
IMPORT_PROJECT_INPUTS_OUTPUTS = 3

print("Opening {}".format(database))
class Worksheet:
    def __init__(self, backend):
        self.backend = backend
        dig = self.backend
        ids = dig.Columns("D")
        #print(ids.Cells(1,1).Value())
        #print(ids.Cells(2,1).Value())
        #print(ids.Cells(3,1).Value())
       

    def get_column_values(self, column):
        return self._filter_out_cells_with_no_values(self.backend.Columns(column).Cells())

    def get_column_values_limited(self, column, max_count):
        #print(column)
        #print("Cell {}".format(self.backend.Columns(column).Cells(1,1).Text))
        cells = self.backend.Columns(column)#.Cells()
        cells_with_values = []
        for i in range(0, max_count-1):
            cells_with_values.append(cells.Cells(i+2,1).Text)
        return cells_with_values
    
    def count_column_values(self, column):
        dig = self.backend
        ids = dig.Columns("D")
        #print(ids.Cells(1,1).Value())
        #print(ids.Cells(2,1).Value())
        #print(ids.Cells(3,1).Value())
        #print(column)
        #print("Cell {}".format(self.backend.Columns(column).Cells(2,1).Text))
        #print("Cell {}".format(self.backend.Columns(column).Cells(3,1).Text))
        values = self._filter_out_row_cells_with_no_values(self.backend.Columns(column).Cells())
        return len(values)

    def get_row_values(self, row):
        dig = self.backend
        ids = dig.Columns(row)
        #print(ids.Cells(1,1).Value())
        #print(ids.Cells(2,1).Value())
        #print(ids.Cells(3,1).Value())
        
        return self._filter_out_row_cells_with_no_values(self.backend.Rows(row).Cells())

    def _filter_out_row_cells_with_no_values(self, cells):
        cells_with_values = []
        for cell in cells:
            if cell is None or cell == "":
                break
            cells_with_values.append(cell)
        return cells_with_values
    
    def _filter_out_cells_with_no_values(self, cells):
        cells_with_values = []
        i = 2
        while True:
            if cells.Cells(i, 0).Value is None or cells.Cells(i, 0).Text == "":
                break
            cells_with_values.append(cells.Cells(i, 0).Text)
            #print("Cell {}".format(cells.Cells(i, 0).Text))
            i+=1
       
        return cells_with_values

    def get_name(self):
        return self.backend.Name

class Workbook:
    def __init__(self, backend):
        self.backend = backend

    def get_worksheets(self):
        worksheets = []
        for worksheet in self.backend.Sheets:
            worksheets.append(Worksheet(worksheet))
        return worksheets
    
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
        self.excel.Interactive = False
        self.excel.DisplayAlerts = False
        self.workbooks = []

    def __enter__(self):
        return self

    def __exit__(self, *args):
        print("Closing workbooks")
        for wb in self.workbooks:
            print("Closing...")
            wb.backend.Close(False)
            del(wb.backend)
        print("Closing excel")
        self.excel.Quit()
        del(self.excel)
        
    def open(self, path):
        try:
           return self._open(path)
        except:
            raise Exception("Could not open workbook: {}".format(path))

    def _open(self, path):
        workbook = Workbook(self.excel.Workbooks.Open(path, None, True))
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

class WorksheetExtractor:
    @staticmethod
    def extract(worksheet, num_entries):
        print("Extracting worksheet", worksheet.get_name())
        top_row = worksheet.get_row_values(1)
        #print(top_row)

        fields = {}

        def _find_data_name_index(data_name):
            try:
                return top_row.index(data_name)
            except:
                raise Exception("Could not find {} in {}".format(data_name, worksheet.get_name()))

        def _get_data_array(data_name):
            data_index = _find_data_name_index(data_name)
            return worksheet.get_column_values_limited(data_index+1, num_entries)
    

        for field in top_row:
            fields[field] = _get_data_array(field)
            #print(fields[field])
        return fields
    
class WorkbookExtractor:
    @staticmethod
    def extract(workbook, exclude_list = []):
        class Entries(object): pass
        entries = {}
        worksheets = workbook.get_worksheets()
        print("Found {} worksheets".format(len(worksheets)))
        for ws in worksheets:
            print(ws.get_name())
        
        for worksheet in worksheets:
            if worksheet.get_name() in exclude_list:
                continue
            columns = worksheet.get_row_values(1)
            index_column = columns.index("database ID") + 1
            entries_count = worksheet.count_column_values(index_column)
            print("Num entries in sheet: {}".format(entries_count))
            entries[worksheet.get_name()] = WorksheetExtractor.extract(worksheet, entries_count)
            #print(entries[worksheet.get_name()]["database ID"])
        #entries.digital_outputs = DigitalOutputExtractor(workbook.get_worksheet_by_name("OutputsDig")).extract_all_digial_outputs()
        return entries
    
class DatabaseEntryExtractorBase:
    def __init__(self, worksheet):
        self.worksheet = worksheet
        print("Extracting worksheet", worksheet.get_name())
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
        entries.worksheets = workbook.get_worksheets()
        print("Found {} worksheets".format(len(entries.worksheets)))
        for ws in entries.worksheets:
            print(ws.get_name())
        entries.parameters = ParameterExtractor(workbook.get_worksheet_by_name("Parameters")).extract_all_parameters()
        #entries.digital_outputs = DigitalOutputExtractor(workbook.get_worksheet_by_name("OutputsDig")).extract_all_digial_outputs()
        return entries

class DatabaseNameFormater:
    @staticmethod
    def format(name):
        return name.replace(" ", "_")

class CommentFormater:
    @staticmethod
    def format(type, sheet, entry):
        to_add = []
        if type is not None:
            to_add.append(("", type))
        if "Comment" in sheet:
            to_add.append(("", sheet["Comment"][entry]))
        for c in ["value", "channel", "pin", "location", "num_dec", "max_value", "min_value", "unit", "max_actual", "min_actual", "unit_actual", "KP", "KI", "frequency", "dither_freq", "dither_value", "byte 3", "byte 2", "byte 1", "byte 0"]:
            if c in sheet:
                to_add.append(("{}:".format(c), sheet[c][entry]))
        
        return ", ".join(["{} {}".format(t, v) for t, v in to_add])
    
class DatabaseIDFormater:
    @staticmethod
    def format_all(sheets):
        return "\n".join("\n\n// {} declarations\n{}".format(k, 
                                                             DatabaseIDFormater.format(k, v, i+1 == len(sheets))
                                                             ) for i, (k, v) in enumerate(sheets.items()))

    @staticmethod
    def format(type, to_format, last):
        return "\n".join(
            "{} := SHL(TO_DWORD(DatabaseType.{}), 16) OR {}{}\t//{}".format(
                DatabaseNameFormater.format(s),
                type,
                i,
                "" if last and i+1 == len(to_format["database ID"]) else ",",
                CommentFormater.format(type, to_format, i)
            ) for i, s in enumerate(to_format["database ID"]))


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
                DatabaseNameFormater.format(s),
                to_format["unit"][i],
                to_format["num_dec"][i],
                ParametersPOUFormater.is_BOOLean(to_format, i),
                CommentFormater.format(None, to_format, i)
            ) for i, s in enumerate(to_format["database ID"]))

    @staticmethod
    def format_definition(to_format):
        return "\n".join(
            "\t{}_{}_{} := Database.get_parameter_as_{}(DatabaseID.{});".format(
                DatabaseNameFormater.format(s),
                to_format["unit"][i],
                to_format["num_dec"][i],
                ParametersPOUFormater.is_BOOLean(to_format, i),
                DatabaseNameFormater.format(s)
            ) for i, s in enumerate(to_format["database ID"]))
    @staticmethod
    def is_BOOLean(to_format, i):
        check = to_format["max_value"][i] + to_format["unit"][i] + to_format["num_dec"][i]
        if check == "1NA0": return "BOOL"
        return "INT"
    
global scale_values
scale_values = ["max_value", "min_value", "value"]
class InitDatabaseFormater:
    @staticmethod
    def format_all(sheets, black_list = ["General"]):
        new_list = {}
        for i, (k, v) in enumerate(sheets.items()):
            if not k in black_list:
                new_list[k] = v
        return "\n".join("\n\n// Initiating {} database\n{}".format(k, 
                                                            InitDatabaseFormater.format(k, v)
                                                             ) for i, (k, v) in enumerate(new_list.items()))

    @staticmethod
    def format(type, to_format):
        return "\n".join(
            "init_{}(database_ID := DatabaseID.{},\t\t{});".format(
                InitDatabaseFormater.pick_method(type),
                DatabaseNameFormater.format(s),
                InitDatabaseFormater.format_init_data(to_format, i)) if not InitDatabaseFormater.should_exclude(to_format, i) else ""
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def pick_method(type):
        if type == "Parameter": return "parameter"
        if type == "AnalogueInput": return "analog"
        if type == "PhysicalDigitalInput": return "digital_input"
        if type == "PhysicalDigitalOutput": return "digital_output"
        if type == "PWMOutput": return "PWM_output"
        return ""
    
    @staticmethod
    def should_exclude(to_format, entry):
        if "Excluded" in to_format and (to_format["Excluded"][entry] == "X" or to_format["Excluded"][entry] == "x"):
            return True
        if "unit_actual" in to_format and to_format["unit_actual"][entry] == "pulse2":
            return True
        return False
    
    @staticmethod
    def format_init_data(to_format, entry, white_list = ["value", "channel", "pin", "location", "num_dec", "max_value", "min_value", "unit", "max_actual", "min_actual", "unit_actual", "KP", "KI", "frequency", "dither_freq", "dither_value"]):
        new_list = {}
        new_list = {}
        for i, (k, v) in enumerate(to_format.items()):
            if k in white_list:
                new_list[k] = v
        return ",\t".join(["{} := {}".format(k, 
                                     InitDatabaseFormater.format_init_value(to_format, k, v, entry)
        ) for i, (k, v) in enumerate(new_list.items())])

    @staticmethod
    def format_init_value(sheet, type, to_format, entry):
        if type == "unit" or type == "unit_actual": return "epirocTypes.Units.{}".format(to_format[entry])
        if type == "location": return "epirocTypes.Location.{}".format(to_format[entry])   
        if "num_dec" in sheet and type in scale_values:
            num_dec = int(sheet["num_dec"][entry])
            to_scale = float(to_format[entry])
            return str(int(to_scale * (10**num_dec)))
        return to_format[entry]
    
global languages
languages = ["Japanese", "Chinese", "Deutsch", "Español", "Français", "Svenska", "Russian", "Italiano", "Norsk", "Polski", "Português", "Suomi"]
class TextlistFormater:
    @staticmethod
    def format_DatabaseID(sheets):
        return [TextlistFormater.format(k, v) for i, (k, v) in enumerate(sheets.items())]

    @staticmethod
    def format(type, to_format):
        return [(TextlistFormater.format_textid(type, i), s, TextlistFormater.format_translations(to_format, i))
            for i, s in enumerate(to_format["database ID"])]
    
    @staticmethod
    def format_textid(type, entry):
        types = ["PhysicalDigitalInput", "PhysicalDigitalOutput", "AnalogueInput", "PWMOutput", "Parameter", "General"]
        #return str("{} << 16 {}".format(types.index(type), entry))
        return str((types.index(type) << 16) + entry)
    
    @staticmethod
    def format_translations(to_format, entry):
        translations = []
        for lan in languages:
            if lan in to_format:
                translations.append((lan, to_format[lan][entry]))
        return translations
    
    @staticmethod
    def format_comment(type, to_format):
        return [(TextlistFormater.format_textid(type, i), s, TextlistFormater.format_translations(to_format, i))
            for i, s in enumerate(to_format["Comment"])]
    

class IOFormater:
    @staticmethod
    def format_analog_input_outputs(to_format):
        return "\n".join(
            "{}_{} : {}; //  {} from database".format(
                DatabaseNameFormater.format(s),
                IOFormater.format_analog_unit_dec(to_format, i),
                "AnalogueInputValue",
                "Analog input")
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def format_analog_unit_dec(to_format, entry):
        return "{}_{}".format(to_format["unit"][entry], to_format["num_dec"][entry])
    
    @staticmethod
    def format_analog_input_var(to_format):
        return "\n".join(
            "{0}_: epirocIOs.{1}(epirocDBD.DatabaseID.{0}, {0}_{2}, IO_reader){3};".format(
                DatabaseNameFormater.format(s),
                IOFormater.select_analog_input_type(to_format["unit_actual"][i]),
                IOFormater.format_analog_unit_dec(to_format, i),
                IOFormater.format_analog_input_var_assignment(to_format, i))
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def select_analog_input_type(type):
        if type == "pulse": return "EncoderInput"
        if type == "pulse3": return "FastcountInput"
        return "AnalogInput"
    
    @staticmethod
    def format_analog_input_var_assignment(to_format, entry):
        assignments = []
        if "Absolute" in to_format and to_format["Absolute"][entry] != "":
            assignments.append("absolute := TRUE")
        elif to_format["unit_actual"][entry] == "pulse":
            assignments.append(
                "retain_value := PersistentVars.{0}_retain, calibrate := InputVariables.{0}_calibrate, calibrate_value := InputVariables.{0}_calibrate_value".format(
                    DatabaseNameFormater.format(to_format["database ID"][entry])))
        if to_format["RigOptions"][entry] != "":
            options = to_format["RigOptions"][entry].split()
            assignments.append("rig_options := {}".format(" OR ".join("epirocConf.RigOptions.{}".format(o) for o in options)))
        if len(assignments) > 0:
            return " := ({})".format(", ".join(a for a in assignments))

        return ""
    
    @staticmethod
    def format_pwm_outputs_input(to_format):
        return "\n".join(
            "{} : {}; // {}".format(
                DatabaseNameFormater.format(s),
                "UINT",
                "Desired output current")
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def format_pwm_outputs_var(to_format):
        return "\n".join(
            "{0}_: epirocIOs.{1}(epirocDBD.DatabaseID.{0}, {0}, IO_reader){2};".format(
                DatabaseNameFormater.format(s),
                "PWMOutput",
                IOFormater.format_pwm_output_var_assignment(to_format, i))
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def format_pwm_output_var_assignment(to_format, entry):
        assignments = []
        if "IdleThresholdfix" in to_format and to_format["IdleThresholdfix"][entry] != "":
            assignments.append("extend_threshold_when_idle := TRUE, extended_threshold := TRAMMING_EXTENDED_PWM_THRESHOLD")
        if "Stuckfix" in to_format and to_format["Stuckfix"][entry] != "":
            assignments.append("fix_stuck_PWM_error := TRUE")
        if len(assignments) > 0:
            return " := ({})".format(", ".join(a for a in assignments))

        return ""
    
    @staticmethod
    def format_digital_input_outputs(to_format):
        return "\n".join(
            "{} : {}; //  {} from database".format(
                DatabaseNameFormater.format(s),
                "epirocUtil.BinaryStateBase",
                "Digital input")
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def format_digital_input_var(to_format):
        return "\n".join(
            "{0}_: DigitalInputBinaryState(epirocDBD.DatabaseID.{0}, {0}.state, IO_reader) := (state := {0});".format(
                DatabaseNameFormater.format(s))
            for i, s in enumerate(to_format["database ID"]))
	
    
    @staticmethod
    def format_digital_output_outputs(to_format):
        return "\n".join(
            "{} : {}; //  {} from database".format(
                DatabaseNameFormater.format(s),
                "BOOL",
                "Digital output")
            for i, s in enumerate(to_format["database ID"]))
    
    @staticmethod
    def format_digital_output_var(to_format):
        return "\n".join(
            "{0}_: epirocIOs.DigitalOutput(epirocDBD.DatabaseID.{0}, {0}, IO_reader);".format(
                DatabaseNameFormater.format(s))
            for i, s in enumerate(to_format["database ID"]))
    
def import_to_library(entries, import_type):
    
    # Import init_database
    print("Import init_database")
    init_database = projects.primary.find("init_database", True)
    if len(init_database) == 0:
        raise Exception("init_database method not found")
    init_database = init_database[0]

    
    init_database.textual_implementation.replace(
    """reset_general_values();

    {}
    """.format(InitDatabaseFormater.format_all(entries)))	

    if import_type == IMPORT_LIB_FLEXI_INTERIM:
        return


    database_ID_enum_text = projects.primary.find("DatabaseID", True)
    if len(database_ID_enum_text) == 0:
        raise Exception("DatabaseID enum not found")
    database_ID_text = None
    if database_ID_enum_text[0].is_textlist:
        if len(database_ID_enum_text) == 1:
            raise Exception("DatabaseID enum not found")
        database_ID_enum = database_ID_enum_text[1]
        database_ID_text = database_ID_enum_text[0]
    else:
        if len(database_ID_enum_text) > 1:
            database_ID_text = database_ID_enum_text[1]
        database_ID_enum = database_ID_enum_text[0]
        
    # Import DatabaseID enum
    print("Import Database IDs")
    database_ID_enum.textual_declaration.replace("""{}
    TYPE DatabaseID :
    (
    {}
    ) DWORD;
    END_TYPE

            """.format("{attribute 'qualified_only'}", DatabaseIDFormater.format_all(entries)))

    # Import Parameters FB
    print("Import parameters")
    paramPOU = POU_Finder.find_POU_by_name(projects.primary, "Parameters")
    if paramPOU is None:
        raise Exception("Parameters POU not found")
    paramPOU.textual_declaration.replace("""PROGRAM Parameters
    VAR_INPUT
        UPDATE : BOOL;
    END_VAR
    VAR_OUTPUT
    {}
    END_VAR
            """.format(ParametersPOUFormater.format_declaration(entries["Parameter"])))

    paramPOU.textual_implementation.replace("""IF UPDATE THEN
    {}
    END_IF
            """.format(ParametersPOUFormater.format_definition(entries["Parameter"])))

    
    if import_type == IMPORT_LIB_NO_TEXTLIST:
        return
    # Look for DatabaseID textlist and delete it and recreate it
    if not database_ID_text is None:
        database_ID_text.remove()

    tlfolder = projects.primary.find("Textlists", True)
    if len(tlfolder) == 0:
        projects.primary.create_folder("Textlists")
        tlfolder = projects.primary.find("Textlists", True)[0]
    else:
        tlfolder = tlfolder[0]
    
    projects.primary.create_textlist("DatabaseID")
    database_ID_enum_text = projects.primary.find("DatabaseID", True)

    if len(database_ID_enum_text) == 0:
        raise Exception("DatabaseID enum not found")
    database_ID_text = None
    if database_ID_enum_text[0].is_textlist:
        if len(database_ID_enum_text) == 1:
            raise Exception("DatabaseID enum not found")
        database_ID_enum = database_ID_enum_text[1]
        database_ID_text = database_ID_enum_text[0]
    else:
        if len(database_ID_enum_text) < 2:
            raise Exception("DatabaseID textlist not found")
        database_ID_enum = database_ID_enum_text[0]
        database_ID_text = database_ID_enum_text[1]

    database_ID_text.move(tlfolder, -1)
    for e in TextlistFormater.format_DatabaseID(entries):
        for (i, id, translations) in e:
            database_ID_text.rows.add(i, id)
            for (lan, tran) in translations:
                database_ID_text.addlanguage(lan)
                database_ID_text.rows[i].setlanguagetext(lan, tran)

    
    def import_comment(textlist, type):
        comment_textlist = projects.primary.find(textlist, True)
        if len(comment_textlist) == 1:
            comment_textlist[0].remove()
        projects.primary.create_textlist(textlist)
        comment_textlist = projects.primary.find(textlist, True)[0]
        comment_textlist.move(tlfolder, -1)
        for (i, id, translations) in TextlistFormater.format_comment(type, entries[type]):
            comment_textlist.rows.add(i, id)
            for (lan, tran) in translations:
                comment_textlist.addlanguage(lan)
                comment_textlist.rows[i].setlanguagetext(lan, tran)
    # Import ParameterDescription            
    import_comment("ParameterDescriptions", "Parameter")
    # Import ParameterDescription            
    import_comment("DatabaseGeneralDescriptions", "General")

def import_to_master(entries):
    # Analog inputs
    analog_input = projects.primary.find("InputsAnalog", True)
    if len(analog_input) == 0:
        raise Exception("InputsAnalog FB not found")
    analog_input = analog_input[0]

    
    analog_input.textual_declaration.replace(
        """FUNCTION_BLOCK InputsAnalog EXTENDS epirocIOS.IOBaseBlock
VAR_INPUT
END_VAR
VAR_OUTPUT
{}
END_VAR
VAR
{}
END_VAR
        """.format(IOFormater.format_analog_input_outputs(entries["AnalogueInput"]),    
                   IOFormater.format_analog_input_var(entries["AnalogueInput"])))

    # PWM Outputs
    pwm_outputs = projects.primary.find("OutputsPWM", True)
    if len(pwm_outputs) > 0:

        pwm_outputs = pwm_outputs[0]

        
        pwm_outputs.textual_declaration.replace(
            """FUNCTION_BLOCK OutputsPWM EXTENDS epirocIOS.IOBaseBlock
VAR CONSTANT
	TRAMMING_EXTENDED_PWM_THRESHOLD: INT := 100;
END_VAR
VAR_INPUT
{}
END_VAR
VAR_OUTPUT
END_VAR
VAR
{}
END_VAR
            """.format(IOFormater.format_pwm_outputs_input(entries["PWMOutput"]),    
                    IOFormater.format_pwm_outputs_var(entries["PWMOutput"])))


        
    #Digital inputs
    digital_input = projects.primary.find("InputsDigital", True)
    if len(digital_input) == 0:
        raise Exception("InputsDigital FB not found")
    digital_input = digital_input[0]

    
    digital_input.textual_declaration.replace(
        """FUNCTION_BLOCK InputsDigital EXTENDS epirocIOS.IOBaseBlock
VAR_INPUT
END_VAR
VAR_OUTPUT
{}
END_VAR
VAR
{}
END_VAR
        """.format(IOFormater.format_digital_input_outputs(entries["PhysicalDigitalInput"]),    
                   IOFormater.format_digital_input_var(entries["PhysicalDigitalInput"])))


    digital_output = projects.primary.find("OutputsDigital", True)
    if len(digital_output) == 0:
        raise Exception("OutputsDigital FB not found")
    digital_output = digital_output[0]

    
    digital_output.textual_declaration.replace(
        """FUNCTION_BLOCK OutputsDigital EXTENDS epirocIOS.IOBaseBlock
VAR_INPUT
{}
END_VAR
VAR
{}
END_VAR
        """.format(IOFormater.format_digital_output_outputs(entries["PhysicalDigitalOutput"]),    
                   IOFormater.format_digital_output_var(entries["PhysicalDigitalOutput"])))


from codesysutil import *
def run():
    try:
        if import_type < 0:
            return
        with Excel() as excel:
            wb = excel.open(database)
            #ws = wb.get_worksheet_by_name("Parameters")
            entries = WorkbookExtractor.extract(wb, exclude_list=["Metadata", "IDs", "RigErrors"])
            for i, (k, v) in enumerate(entries.items()):
                print("{}, Entries: {}, {}".format(k, len(v), len(v["database ID"])))
                #for e in v["database ID"]:
                # print(e)

            #database_ID_enum = POU_Finder.find_POU_by_name(projects.primary, "DatabaseID")
            if import_type < IMPORT_PROJECT_INPUTS_OUTPUTS:
                import_to_library(entries, import_type)
            elif import_type == IMPORT_PROJECT_INPUTS_OUTPUTS:
                import_to_master(entries)
    except:
        raise
    finally:
        print("--- Script finished. ---")
run()