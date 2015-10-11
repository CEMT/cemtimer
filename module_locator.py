import sys
import os
#Credit to Daniel Stutzbach's answer: http://stackoverflow.com/a/2632297/5221035
#MUST include the module.py file in the build folder - this is for the Windows sched task
def we_are_frozen():
    # All of the modules are built-in to the interpreter, e.g., by py2exe
    return hasattr(sys, "frozen")

def module_path():
    encoding = sys.getfilesystemencoding()
    if we_are_frozen():
        return os.path.dirname(unicode(sys.executable, encoding))
    return os.path.dirname(unicode(__file__, encoding))