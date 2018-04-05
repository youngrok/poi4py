import os

import jpype
from jpype import java


classloader = None


def start_jvm():
    global classloader

    if jpype.isJVMStarted():
        if not jpype.isThreadAttachedToJVM():
            jpype.attachThreadToJVM()
            java.lang.Thread.currentThread().setContextClassLoader(classloader)
        return

    java_library_path = f'{os.path.dirname(__file__)}/java/'
    classpath = os.pathsep.join(f'{java_library_path}/{lib}' for lib in [
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
    ])

    jpype.startJVM(jpype.get_default_jvm_path(),
                   f'-Djava.class.path={classpath}',
                   '-Dfile.encoding=UTF8',
                   '-ea',
                   '-Xmx1024m')

    classloader = java.lang.Thread.currentThread().getContextClassLoader()


def open_workbook(filename, password=None):
    try:
        start_jvm()
        return jpype.JPackage('org').apache.poi.ss.usermodel.WorkbookFactory.create(java.io.File(filename), password)
    except jpype.JavaException as e:
        print(e.message())
        print(e.stacktrace())
        raise e


def shutdown_jvm():
    jpype.shutdownJVM()
