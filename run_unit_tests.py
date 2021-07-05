# encoding:utf-8
# We enable the new python 3 print syntax
from __future__ import print_function
from scriptengine import *

class TestCase:
    def __init__(self, test_POU, prio = 100):
        self.test_POU = test_POU
        self.prio = prio

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
    def find_all_test_case_POUs(project):
        POUs = []
        for child in project.get_children():
            POUs_to_append = POU_Finder.find_all_test_case_POUs_in_tree(child)
            for to_append in POUs_to_append:
                POUs.append(to_append)
        return POUs

    @staticmethod
    def find_all_test_case_POUs_in_tree(treeobj):
        POUs = []
        
        if treeobj.has_textual_declaration:
            textual_declaration = treeobj.textual_declaration
            first_line = textual_declaration.get_line(0)
            if first_line.find("(*Test") >= 0 and first_line.find("ABSTRACT") == -1:
                PRIO_pos = first_line.find("PRIO := ")
                if PRIO_pos >= 0:
                    prio = int(first_line[len("PRIO := ") + PRIO_pos])
                    POUs.append(TestCase(treeobj, prio))
                else:
                    POUs.append(TestCase(treeobj))
        else:
            for child in treeobj.get_children(False):
                POUs_to_append = POU_Finder.find_all_test_case_POUs_in_tree(child)
                for to_append in POUs_to_append:
                    POUs.append(to_append)
        
        return POUs

class CodesysTypeConverter:
    @staticmethod
    def to_int(codesys_int):
        return int(codesys_int.split("#")[1])

class TestCaseSorter:
    @staticmethod
    def sort_by_prio(test_cases):
        test_cases.sort(key = lambda test_case: test_case.prio)



plc_prg = POU_Finder.find_POU_by_name(projects.primary, "PLC_PRG")
test_cases = POU_Finder.find_all_test_case_POUs(projects.primary)
if len(test_cases) == 0:
    raise Exception("No test cases found.")

test_cases.reverse()
TestCaseSorter.sort_by_prio(test_cases)

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
            "\n\t".join("test_{0} : {1} := (test_case_name := '{1}');".format(index, test_case.test_POU.get_name(False)) for index, test_case in enumerate(test_cases)), 
            ",\n\t\t".join("test_{0}".format(index) for index, test_case in enumerate(test_cases))))

projects.primary.save()

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