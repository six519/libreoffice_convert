#!/usr/bin/env python3
from .converter import PythonLibreOffice
import sys

def showErrorMessage(msg):
    sys.stderr.write(msg)

def libreoffice_convert():

    if len(sys.argv) == 3:
        conv = PythonLibreOffice()
        conv.convertFile(sys.argv[1], sys.argv[2])

        if conv.lastError != "":
            showErrorMessage("\nAn error occurred while converting the file: %s\n" % conv.lastError)
    else:
        showErrorMessage("\nUsage: libreoffice_convert OUTPUT_EXTENSION INPUT_FILE\n")