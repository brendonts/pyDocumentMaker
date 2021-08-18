# This Python3 script takes a formatted Excel spreadsheet from the dir of the .py script and creates multiple doc documents using the provided Word template in the same dir.

# Import dependencies, docx-mailmerge, and pandas are required
import os
from typing import Protocol
import pandas as pd
import mailmerge
from mailmerge import MailMerge

# Specify current dir for soruce files
dir = os.path.abspath('')
files = os.listdir(dir)
print(dir)
template = "doc-autoTemplate.docx"

# Takes the xlsx containing doc info and converts in to Pandas data frame for parsing
df = pd.DataFrame()
for file in files:
        if file.endswith('.xlsx'):
                df = df.append(pd.read_excel(file), ignore_index=True)
df.head()
df.to_excel('doc_Output.xlsx')

# Create doc Class, initialize all properties. Note, each form field must be a property which can become quite cumbersome.
class doc:
    def __init__(self, docName, vendor, application, service, ports, protocol, boundary, comments, description, implementation, attribute, tpocDSN, browserBased, clientServer, peerToPeer, networkProtocol, currentDate, syspocName, syspocTitle, syspocOrg, syspocEmail, tpocName, tpocTitle, tpocEmail, tpocOrg, ppsPOCname, ppsPOCtitle, ppsPOCorg, ppsPOCDSN, ppsPOCemail):
        self.docName = docName
        self.vendor = vendor
        self.application = application
        self.service = service
        self.ports = ports
        self.protocol = protocol
        self.boundary = boundary
        self.comments = comments
        self.description = description
        self.implementation = implementation
        self.attribute = attribute
        self.tpocDSN = tpocDSN
        self.browserBased = browserBased
        self.clientServer = clientServer
        self.peerToPeer = peerToPeer
        self.networkProtocol = networkProtocol
        self.currentDate = currentDate
        self.syspocName = syspocName
        self.syspocTitle = syspocTitle
        self.syspocOrg = syspocOrg
        self.syspocEmail = syspocEmail
        self.tpocName = tpocName
        self.tpocTitle = tpocTitle
        self.tpocEmail = tpocEmail
        self.tpocOrg = tpocOrg
        self.ppsPOCname = ppsPOCname
        self.ppsPOCtitle = ppsPOCtitle
        self.ppsPOCorg = ppsPOCorg
        self.ppsPOCDSN = ppsPOCDSN
        self.ppsPOCemail = ppsPOCemail
    
    # Method for creating doc form using the template from MailMerge
    def createdoc(self):
        document = MailMerge('doc-autoTemplate.docx') # Open the doc template from current dir
        print(document.get_merge_fields())
        document.merge(
            doctitle=self.docName,
            vendor=self.vendor,
            doc=self.comments,
            applicationName=self.application,
            dataServiceShortName=self.service,
            ports=str(self.ports),
            protocol=self.protocol,
            boundary=self.boundary,
            comments=self.comments,
            serviceDescription=self.description,
            implementationInformation=self.implementation,
            attribute=self.attribute,
            #tpocDSN=self.tpocDSN,
            browserBased=self.browserBased,
            clientServer=self.clientServer,
            peerToPeer=self.peerToPeer,
            networkProtocol=self.networkProtocol,
            #currentDate=self.currentDate, # Need to fix the date field erroring out
            syspocName=self.syspocTitle,
            syspocTitle=self.syspocTitle,
            syspocOrg=self.syspocOrg,
            syspocNIPRemail=self.syspocEmail,
            tpocName=self.tpocName,
            tpocTitle=self.tpocTitle,
            tpocNIPRemail=self.tpocEmail,
            tpocOrg=self.tpocOrg,
            ppspocName=self.ppsPOCname,
            ppspocTitle=self.ppsPOCtitle,
            ppspocOrg=self.ppsPOCorg,
            #ppspocDSN=self.ppsPOCDSN,
            ppspocNIPRemail=self.ppsPOCemail)
        document.write('doc' + " " + self.docName + '.docx')

# Iterate through each row in the pandas dataframe containing doc info
for index, row in df.iterrows():
    docDocName = df.loc[index, 'docDocName']
    vendor = df.loc[index, 'Vendor']
    Application = df.loc[index, 'Application']
    Ports = df.loc[index, 'Ports']
    Protocols = df.loc[index, 'Protocols']
    Boundary = df.loc[index, 'Boundary']
    serviceDescription = df.loc[index, 'serviceDescription']
    implementationDescription = df.loc[index, 'implementationDescription']
    tpocDSN = df.loc[index, 'tpocDSN']
    browserBased = df.loc[index, 'browserBased']
    clientServer = df.loc[index, 'clientServer']
    peerToPeer = df.loc[index, 'peerToPeer']
    networkProtocol = df.loc[index, 'networkProtocol']
    currentDate = df.loc[index, 'currentDate']
    syspocName = df.loc[index, 'syspocName']
    syspocTitle = df.loc[index, 'syspocTitle']
    syspocOrg = df.loc[index, 'syspocOrg']
    syspocEmail = df.loc[index, 'syspocEmail']
    tpocName = df.loc[index, 'tpocName']
    tpocTitle = df.loc[index, 'tpocTitle']
    tpocEmail = df.loc[index, 'tpocEmail']
    tpocOrg = df.loc[index, 'tpocOrg']
    ppsPOCname = df.loc[index, 'ppsPOCname']
    ppsPOCtitle = df.loc[index, 'ppsPOCtitle']
    ppsPOCorg = df.loc[index, 'ppsPOCorg']
    ppsPOCDSN = df.loc[index, 'ppsPOCDSN']
    ppsPOCemail = df.loc[index, 'ppsPOCemail']

    # Create an instance of a doc for the current row in the datafram
    docEntry = doc(docDocName, vendor, Application, "placeHolder", Ports, Protocols, Boundary, "placeHolder", serviceDescription, implementationDescription, "placeHolder", tpocDSN, browserBased, clientServer, peerToPeer, networkProtocol, currentDate, syspocName, syspocTitle, syspocOrg, syspocEmail, tpocName, tpocTitle, tpocEmail, tpocOrg, ppsPOCname, ppsPOCtitle, ppsPOCorg, ppsPOCDSN, ppsPOCemail)
    docEntry.createdoc() # Calls the createdoc method from doc to create a doc Word document using the current row
    print(docDocName)
    print("----- END OF doc ENTRY -----")
