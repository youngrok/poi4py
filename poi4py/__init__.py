import os

import jpype
import logging
from jpype import java, JPackage


def start_jvm():

    if not jpype.isJVMStarted():
        jpype.addClassPath(f'{os.path.dirname(__file__)}/java/*')
        jpype.startJVM(jpype.get_default_jvm_path(),
                       '-Dfile.encoding=UTF8',
                       '-ea',
                       '-Xmx1024m')

    if not jpype.isThreadAttachedToJVM():
        jpype.attachThreadToJVM()


def open_workbook(filename, password=None, read_only=False):
    try:
        start_jvm()
        return jpype.JPackage('org').apache.poi.ss.usermodel.WorkbookFactory.create(
            java.io.File(filename), password, read_only)
    except java.lang.Exception as e:
        logging.exception(e)
        raise e


def shutdown_jvm():
    jpype.shutdownJVM()
