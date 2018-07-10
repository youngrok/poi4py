import os

import jpype
import logging
from jpype import java, JPackage
from jpype.reflect import getDeclaredMethods


classpath_added = False

def start_jvm():

    if not jpype.isJVMStarted():
        jpype.startJVM(jpype.get_default_jvm_path(),
                       '-Dfile.encoding=UTF8',
                       '-ea',
                       '-Xmx1024m')

    if not jpype.isThreadAttachedToJVM():
        jpype.attachThreadToJVM()

    global classpath_added
    if not classpath_added:
        add_classpaths([f'{os.path.dirname(__file__)}/java/{lib}' for lib in [
            'poi-3.17.jar',
            'poi-excelant-3.17.jar',
            'poi-ooxml-3.17.jar',
            'poi-ooxml-schemas-3.17.jar',
            'poi-scratchpad-3.17.jar',
            'lib/commons-codec-1.10.jar',
            'lib/commons-collections4-4.1.jar',
            'lib/commons-logging-1.2.jar',
            'lib/log4j-1.2.17.jar',
            'ooxml-lib/xmlbeans-2.6.0.jar',
            'ooxml-lib/curvesapi-1.04.jar',
        ]])
        classpath_added = True


def add_classpaths(paths):
    URL = JPackage('java.net').URL
    sysloader = JPackage('java.lang').ClassLoader.getSystemClassLoader()
    method = [m for m in (getDeclaredMethods(JPackage('java.net').URLClassLoader)) if m.name == 'addURL'][0]
    method.setAccessible(True)
    for path in paths:
        method.invoke(sysloader, [URL(f"file://{path}")])


def open_workbook(filename, password=None, read_only=False):
    try:
        start_jvm()
        return jpype.JPackage('org').apache.poi.ss.usermodel.WorkbookFactory.create(
            java.io.File(filename), password, read_only)
    except jpype.JavaException as e:
        logging.exception(e)
        raise e


def shutdown_jvm():
    jpype.shutdownJVM()
