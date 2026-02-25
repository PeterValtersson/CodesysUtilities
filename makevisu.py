#!/usr/local/bin/python
# -*- coding: utf-8 -*-
from __future__ import print_function
from scriptengine import *


from codesysutil import *

def findApplication(proj):
    apps = proj.find("Application", True)
    if len(apps) != 1:
        raise Exception("Unexpected number of apps: " + str(len(apps)))
    return apps[0]

def findVisualization(proj, visuName):
    visus = proj.find(visuName, True)
    if  len(visus) == 0:
        return None
    elif len(visus) != 1:
        if visus[0].is_visualobject:
            return visus[0]
        elif visus[1].is_visualobject:
            return visus[1]
        raise Exception("Unexpected number of visus: " + str(len(visus)))
    return visus[0]

elements2 = ["Button", 
            "Label", 
            "Numberbox", 
            "CheckBox", 
            "Dropdown", 
            "Groupbox", 
            "Image", 
            "Gauge", 
            "Stringbox", 
            "AngleIndicationElement", 
            "DrillmetersElement", 
            "LineElement", 
            "Scrollbar"]
elements = ["Rectangle",
            "Image",
            "Points",
            "Ellipse"]
def make_frame(visu, index):
    refs = []
    for element in elements:
        visuRef = visu.create_frame_reference(element)
        visuRef.set_parameter("element", "container.get_element({0})".format(index))
        refs.append(visuRef)
        
    frame = visu.visual_element_list.add_element(VisualElementType.Frame)

    frame.set_frame_references(refs)
    return frame

def clear_elements(elementList):
    for i in range(0, len(elementList)):
        elementList.remove_at(0)

proj = projects.primary
#app = findApplication(proj)

visu = findVisualization(proj, "ElementContainer")
if visu != None:
    visu.begin_modify()
    clear_elements(visu.visual_element_list)

    for i in range(0, 1000):
        frame = make_frame(visu, i)
        frame.set_property("Swiping preview", False)   
        frame.set_property("Switch frame variable.Variable", "container.get_element({0}).get_fragment_type()".format(i))
        frame.set_property("Scaling type", "Fixed")    
        frame.set_property("State variables.Invisible", "NOT container.is_frame_visible({0})".format(i))

    for i in range(0, 4):
        visuRef1 = visu.create_frame_reference("ElementContainerClipped")
        visuRef1.set_parameter("container", "container.get_container({0})".format(i))
        ewElem = visu.visual_element_list.add_element(VisualElementType.Frame)
        ewElem.set_property("Scaling type", "Fixed")    
        #ewElem.set_property("Swiping preview", False)    
        ewElem.set_property("Clipping", True)    
        ewElem.set_property("Show frame", "No frame")    
        ewElem.set_property("State variables.Invisible", "NOT container.is_container_frame_visible({0})".format(i))
        ewElem.set_property("Relative movement.Movement top-left.X", "container.get_container({0}).startX".format(i))  
        ewElem.set_property("Relative movement.Movement top-left.Y", "container.get_container({0}).startY".format(i))  
        ewElem.set_property("Relative movement.Movement bottom-right.X", "container.get_container({0}).endX-150".format(i))  
        ewElem.set_property("Relative movement.Movement bottom-right.Y", "container.get_container({0}).endY-30".format(i))  
        
        #visuRef3.set_parameter("pou", "PLC_PRG.inst1")
        ewElem.set_frame_references([visuRef1])


    visu.end_modify()
    visu = findVisualization(proj, "ElementContainerClipped")
    visu.begin_modify()
    #newElem = elementList.add_element(VisualElementType.Rectangle)
    #newElem.set_property("Position.X", 5)
    #newElem.set_property("Position.Y", 5)
    clear_elements(visu.visual_element_list)


    for i in range(0, 99):
        frame = make_frame(visu, i)
        frame.set_property("Switch frame variable.Variable", "container.get_element({0}).get_fragment_type()".format(i))
        frame.set_property("Scaling type", "Fixed")    
        frame.set_property("Swiping preview", False)    
        frame.set_property("State variables.Invisible", "NOT container.is_frame_visible({0})".format(i))
        #frame.set_property("Relative movement.Movement top-left.X", "-container.startX")  
        #frame.set_property("Relative movement.Movement top-left.Y", "-container.startY")  
        #frame.set_property("Relative movement.Movement bottom-right.X", "-container.startX")  
        #frame.set_property("Relative movement.Movement bottom-right.Y", "-container.startY")
        
    
    visu.end_modify()

visu = findVisualization(proj, "Visualization")
if visu != None:
    visu.begin_modify()
    clear_elements(visu.visual_element_list)
    visuRef1 = visu.create_frame_reference("epirocDisp.ElementContainer")
    visuRef1.set_parameter("container", "POU.pou.container")
    ewElem = visu.visual_element_list.add_element(VisualElementType.Frame)
    ewElem.set_frame_references([visuRef1])
    ewElem.set_property("Scaling type", "Fixed")    
    ewElem.set_property("Clipping", False)    
    ewElem.set_property("Show frame", "No frame")    
    for i in range(0, 3):
        visuRef1 = visu.create_frame_reference("epirocDisp.ElementContainerClipped")
        visuRef1.set_parameter("container", "POU.pou.container.get_container({0})".format(i))
        ewElem = visu.visual_element_list.add_element(VisualElementType.Frame)
        ewElem.set_property("Scaling type", "Fixed")    
        ewElem.set_property("Clipping", True)    
        ewElem.set_property("Show frame", "No frame")    
        ewElem.set_property("State variables.Invisible", "NOT POU.pou.container.is_container_frame_visible({0})".format(i))
        ewElem.set_property("Relative movement.Movement top-left.X", "POU.pou.container.get_container({0}).startX".format(i))  
        ewElem.set_property("Relative movement.Movement top-left.Y", "POU.pou.container.get_container({0}).startY".format(i))  
        ewElem.set_property("Relative movement.Movement bottom-right.X", "POU.pou.container.get_container({0}).endX-150".format(i))  
        ewElem.set_property("Relative movement.Movement bottom-right.Y", "POU.pou.container.get_container({0}).endY-30".format(i))  
        
        #visuRef3.set_parameter("pou", "PLC_PRG.inst1")
        ewElem.set_frame_references([visuRef1])


    visu.end_modify()