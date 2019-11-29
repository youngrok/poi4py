import os
import subprocess

import jpype
import logging

import sys
from jpype import java, JPackage


def start_jvm(jvm_path=None):
    if not jvm_path:
        if sys.platform == 'darwin':
            os.environ['JAVA_HOME'] = subprocess.check_output(['/usr/libexec/java_home']).strip().decode('utf8')

        jvm_path = jpype.getDefaultJVMPath()

    if not jpype.isJVMStarted():
        jpype.addClassPath(f'{os.path.dirname(__file__)}/java/*')
        jpype.startJVM(jvm_path,
                       '-Dfile.encoding=UTF8',
                       '-ea',
                       '-Xmx1024m',
                       convertStrings=False)

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
