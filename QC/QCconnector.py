'''
@author: Michel Lalmohamed
'''
# import the pypiwin32 package locally (pip)
from win32com.client import Dispatch

class QC_ConnectorClass(object):

    '''
    Connector Class which returns as the connection object for further QC API use
    '''
    def __init__(self):
        print("class init")
    # This function creates a connection to the HP QC server and logs in to the server by using the
    # user credentials

    def connectToQC(self, qcServer, qcUser, qcPassword, qcDomain, qcProject):
        #HP QC OTA methods
        self.TD = Dispatch("TDApiOle80.TDConnection.1")
        self.TD.InitConnectionEx(qcServer)
        self.TD.Login(qcUser,qcPassword)
        self.TD.Connect(qcDomain,qcProject)
        print("Logged in")
        # Return the connection object
        return self.TD