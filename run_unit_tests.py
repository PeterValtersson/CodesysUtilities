# encoding:utf-8
# We enable the new python 3 print syntax
from __future__ import print_function
from scriptengine import *

class POU_Finder:
    @staticmethod
    def find_POU_by_name(project, POU_name):
        for obj in project.get_children():
            plc_prg = POU_Finder.find_POU_in_tree_by_name(obj, POU_name)
            if not plc_prg is None:
                return plc_prg

        raise Exception("Could not find PLC_PRG")

    @staticmethod
    def find_POU_in_tree_by_name(treeobj, POU_name):
        name = treeobj.get_name(False)
        if name == POU_name:
            return treeobj

        for child in treeobj.get_children(False):
            plc_prg =  POU_Finder.find_POU_in_tree_by_name(child, POU_name)
            if not plc_prg is None:
                return plc_prg
                
    @staticmethod
    def find_all_POUs_extending_from(project, POU_name):
        POUs = []
        for child in project.get_children():
            POUs_to_append = POU_Finder.find_all_POUs_in_tree_extending_from(child, POU_name)
            for to_append in POUs_to_append:
                POUs.append(to_append)
        return POUs

    @staticmethod
    def find_all_POUs_in_tree_extending_from(treeobj, POU_name):
        POUs = []
        
        if treeobj.has_textual_declaration:
            textual_declaration = treeobj.textual_declaration
            first_line = textual_declaration.get_line(0)
            if first_line.find("EXTENDS {}".format(POU_name)) > 0:
                POUs.append(treeobj)
        else:
            for child in treeobj.get_children(False):
                POUs_to_append = POU_Finder.find_all_POUs_in_tree_extending_from(child, POU_name)
                for to_append in POUs_to_append:
                    POUs.append(to_append)
        
        return POUs

class CodesysTypeConverter:
    @staticmethod
    def to_int(codesys_int):
        return int(codesys_int.split("#")[1])

# Define the printing function. This function starts with the
# so called "docstring" which is the recommended way to document
# functions in python.
def print_tree(treeobj, depth=0):
    """ Print a device and all its children

    Arguments:
    treeobj -- the object to print
    depth -- The current depth within the tree (default 0).

    The argument 'depth' is used by recursive call and
    should not be supplied by the user.
    """

    # if the current object is a device, we print the name and device identification.
    if treeobj.has_textual_declaration:
        textual_declaration = treeobj.textual_declaration
        first_line = textual_declaration.get_line(0)
        if first_line.find("EXTENDS TestCaseBase") > 0:
            name = treeobj.get_name(False)
            print("{0}- {1}".format("--"*depth, name))

    # we recursively call the print_tree function for the child objects.
    for child in treeobj.get_children(False):
        print_tree(child, depth+1)

plc_prg = POU_Finder.find_POU_by_name(projects.primary, "PLC_PRG")
test_cases = POU_Finder.find_all_POUs_extending_from(projects.primary, "TestCaseBase")
if len(test_cases) == 0:
    raise Exception("No test cases found.")

test_cases.reverse()

plc_prg.textual_declaration.replace(
"""PROGRAM PLC_PRG
VAR CONSTANT
    NUM_TESTS : INT := {0};
END_VAR
VAR
    test_index : INT;
    {1}
    tests : ARRAY[0..NUM_TESTS-1] OF TestCase := 
    [
        {2}
    ];
END_VAR
""".format( len(test_cases), 
            "\n\t".join("test_{0} : {1} := (test_case_name := '{1}');".format(index, test_case.get_name(False)) for index, test_case in enumerate(test_cases)), 
            ",\n\t\t".join("test_{0}".format(index) for index, test_case in enumerate(test_cases))))

application = projects.primary.active_application
application.build()

online_application = online.create_online_application(application)

# login to application.
online_application.login(OnlineChangeOption.Never, True)
#online_application.set_prepared_value("PLC_PRG.test_index", "0")
#online_application.write_prepared_values()
#online_application.reset(reset_option = ResetOption.Cold, force_kill=True)

#while online_application.application_state != ApplicationState.stop:
#    wait = True

# start PLC if necessary
if online_application.application_state != ApplicationState.run:
    online_application.start()


#test_index = CodesysTypeConverter.to_int(online_application.read_value("PLC_PRG.test_index"))
#while test_index < len(test_cases):
#    print("Running test: {0}".format(test_cases[test_index]))
#    test_index = CodesysTypeConverter.to_int(online_application.read_value("PLC_PRG.test_index"))
    

print("--- Script finished. ---")