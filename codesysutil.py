
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

        raise Exception("Could not find {}".format(POU_name))

    @staticmethod
    def find_POU_in_tree_by_name(treeobj, POU_name):
        name = treeobj.get_name(False)
        if name == POU_name and treeobj.has_textual_declaration:
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
                blank_split = first_line.split()
                prio = 0
                for i, v in enumerate(blank_split):
                    if v == "PRIO":
                        print(blank_split)
                        print(blank_split[i+2])
                        print(filter(str.isdigit, blank_split[i+2]))
                        prio = int(filter(str.isdigit, blank_split[i+2]))
                POUs.append(TestCase(treeobj, prio))
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
