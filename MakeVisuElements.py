from scriptengine import *

res = system.ui.query_string("Visualization Name", multi_line=False)
if (res):
    objects = projects.primary.find(res, True)
    for o in objects:
        print(o.get_name())
        try:
            print(o.is_visualobject)
        except:
            print('Not visu')

    print(objects[0])
    print(len(objects))