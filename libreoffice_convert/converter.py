"""
    LIBRE OFFICE CONVERT
    ====================

    * Python module using LibreOffice API to convert file format to another file format

    THANKS
    ======

    * Thanks to: Mirko Nasato for his PyODConverter http://www.artofsolving.com/opensource/pyodconverter

    TESTED USING
    ============

    * Fedora release 20 (Heisenbug)
    * Python 3.3.2

    INSTALL DEPENDENCIES
    ====================

    * yum install libreoffice-sdk

"""

import uno
import subprocess
import time
import os

from com.sun.star.beans import PropertyValue
from .document_types import *

LIBREOFFICE_DEFAULT_PORT = 6519
LIBREOFFICE_DEFAULT_HOST = "localhost"

class PythonLibreOffice(object):

    def __init__(self, host=LIBREOFFICE_DEFAULT_HOST, port=LIBREOFFICE_DEFAULT_PORT):
        self.host = host
        self.port = port
        self.local_context = uno.getComponentContext()
        self.resolver = self.local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", self.local_context)
        self.connectionString = "socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % (LIBREOFFICE_DEFAULT_HOST, LIBREOFFICE_DEFAULT_PORT)
        self.context = None
        self.desktop = None
        self.runUnoProcess()
        self.__lastErrorMessage = ""

        try:
            self.context = self.resolver.resolve("uno:%s" % self.connectionString)
            self.desktop = self.context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)
        except Exception as e:
            self.__lastErrorMessage = str(e)

    @property 
    def lastError(self):

        return self.__lastErrorMessage

    def terminateProcess(self):

        try:
            if self.desktop:
                self.desktop.terminate()
        except Exception as e:
            self.__lastErrorMessage = str(e)
            return False

        return True

    def convertFile(self, outputFormat, inputFilename):

        if self.desktop:
        
            tOldFileName = os.path.splitext(inputFilename)
            outputFilename = "%s.%s" % (tOldFileName[0], outputFormat)
            inputFormat = tOldFileName[1].replace(".","")
            inputUrl = uno.systemPathToFileUrl(os.path.abspath(inputFilename))
            outputUrl = uno.systemPathToFileUrl(os.path.abspath(outputFilename))

            if inputFormat in LIBREOFFICE_IMPORT_TYPES:
                inputProperties = {
                    "Hidden": True
                }

                inputProperties.update(LIBREOFFICE_IMPORT_TYPES[inputFormat])

                doc = self.desktop.loadComponentFromURL(inputUrl, "_blank", 0, self.propertyTuple(inputProperties))
                
                try:
                    doc.refresh()
                except:
                    pass

                docFamily = self.getDocumentFamily(doc)
                if docFamily:
                    try:
                        outputProperties = LIBREOFFICE_EXPORT_TYPES[outputFormat][docFamily]
                        doc.storeToURL(outputUrl, self.propertyTuple(outputProperties))
                        doc.close(True)

                        return True
                    except Exception as e:
                        self.__lastErrorMessage = str(e)
        
        self.terminateProcess()

        return False

    def propertyTuple(self, propDict):
        properties = []
        for k,v in propDict.items():
            property = PropertyValue()
            property.Name = k
            property.Value = v
            properties.append(property)

        return tuple(properties)

    def getDocumentFamily(self, doc):
        try:
            if doc.supportsService("com.sun.star.text.GenericTextDocument"):
                return LIBREOFFICE_DOC_FAMILIES[0]
            if doc.supportsService("com.sun.star.text.WebDocument"):
                return LIBREOFFICE_DOC_FAMILIES[1]
            if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                return LIBREOFFICE_DOC_FAMILIES[2]
            if doc.supportsService("com.sun.star.presentation.PresentationDocument"):
                return LIBREOFFICE_DOC_FAMILIES[3]
            if doc.supportsService("com.sun.star.drawing.DrawingDocument"):
                return LIBREOFFICE_DOC_FAMILIES[4]
        except:
            pass

        return None

    def runUnoProcess(self):
        subprocess.Popen('soffice --headless --norestore --accept="%s"' % self.connectionString, shell=True, stdin=None, stdout=None, stderr=None)
        time.sleep(3)