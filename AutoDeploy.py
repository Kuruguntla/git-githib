from ftplib import FTP
import xml.etree.ElementTree as ET
import fileinput
import yaml
import os
import base64
import shutil
import subprocess
import zipfile
import logging,sys
import ftputil
import time
import socket


if(os.name=="nt"):
    tempStorageLocation = "c://SDM_Patch"
else:
    tempStorageLocation = "//SDM_Patch"

if not os.path.exists(tempStorageLocation):
    os.makedirs(tempStorageLocation)

logging.basicConfig(filename=tempStorageLocation + "//AutoDeploy.log",level=logging.INFO,filemode="w",format='%(asctime)s:%(levelname)s:%(message)s')

logger = logging.getLogger()
sys.stderr.write = logger.error
sys.stdout.write = logger.info

# Start APPLYPTF related functions
def patchExtraction(patchDestination,commonInstallerDestination,updateFile):
    os.chdir(patchDestination)
    currentFolderList = os.listdir(patchDestination)
    for eachFile in currentFolderList:
        if(updateFile!=eachFile):
            logging.info("Removing unwanted file: " + eachFile)
            os.remove(eachFile)
    for eachFile in currentFolderList:
        if(updateFile in eachFile):
            logging.info("Found Patch: " + eachFile)
            logging.info("Extracting patch " + eachFile + " under " + patchDestination)
            command = commonInstallerDestination + "//filestore//utils//CAZIP//cazipxp.exe -u " + eachFile
            logging.info(command)
            p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
            stdout, stderr = p.communicate()
            logging.info("Caz file extraction completed with exit code " + str(p.returncode))
    #Rename caz file if different from JCL
    currentFolderList = os.listdir(patchDestination)
    for eachFile in currentFolderList:
        if (".JCL" in eachFile):
            logging.info("Current JCL file is : " + eachFile)
            if(updateFile[:-4] != eachFile[:-4]):
                logging.info("Renaming " + updateFile + " to " + eachFile[:-4] + ".caz")
                os.rename(updateFile,eachFile[:-4] + ".caz")
            if(applyPatch(eachFile,patchDestination,commonInstallerDestination)==60):
                logging.info("applyPatch function succeeded with exit code 60")
            else:
                logging.info("applyPatch function failed" )
    return

def applyPatch(jclFileName,patchDestination,commonInstallerDestination):
    command = commonInstallerDestination + "\\filestore\\utils\\ApplyPTF\\APPLYPTF /SILENT /OUTPUTFILE=" + patchDestination + "\\" + jclFileName[:-4] + "_Binary_output.log" + " /DEBUGFILE=" + patchDestination + "\\" + jclFileName[:-4] + "_Binary_debug_.log" +  " /PTF=" + patchDestination + "\\" + jclFileName + " /INSTALLAGAIN /INSTALLNEW  /NODE=" + socket.gethostname()
    logging.info("ApplyPTF command to be executed : " + command)
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("ApplyPTF with exit code " + str(p.returncode))
    return p.returncode

def copyWebXmlFiles(fileNameToCopy,source,dest):
    logging.info("Copying " + fileNameToCopy + " from " + source + " to " + dest)
    shutil.copyfile(source,dest)
    return

def applyDatFiles(sdmPatchDestination,sdmInstalledLocation,stringToMatch):
    os.chdir(sdmPatchDestination)
    currentFolderList = os.listdir(sdmPatchDestination)
    for eachDatFile in currentFolderList:
        if ((stringToMatch) in eachDatFile):
            logging.info("Found Dat file : " + eachDatFile)
            logging.info("Need to apply the Dat File " + eachDatFile)
            command = "pdm_load -f " + sdmInstalledLocation + "\\data\\tagged\\" + eachDatFile + " -v"
            logging.info("Executing Dat update command :" + command)
            p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
            stdout, stderr = p.communicate()
            logging.info("PDM_load completed with exit code " + str(p.returncode))
    return
# End ApplyPTF related functions

# Create SDM Folder structure
def createFolderStructure(tempStorageLocation, folderVersion, productUpdate, commonInstallerDestination):


    folderToRemove=os.listdir(commonInstallerDestination + "//patches//" + productUpdate)
    for folder in folderToRemove:
        if(folder!=folderVersion):
            shutil.rmtree(commonInstallerDestination +"//patches//" + productUpdate + "//" + folder)

    if(os.path.exists(tempStorageLocation)):
        folderList = os.listdir(tempStorageLocation)
        for folder in folderList:
            if(folder == productUpdate) and ("xFlow" not in productUpdate):
                currentFolderVersion = os.listdir(tempStorageLocation + "//" + folder)
                for eachFolder in currentFolderVersion:
                    shutil.copytree(tempStorageLocation + "//" + productUpdate + "//" + eachFolder,commonInstallerDestination + "//patches//" + productUpdate + "//" + eachFolder)
            if("xFlow" in productUpdate):
                if(folder!=folderVersion):
                    shutil.copytree(tempStorageLocation + "//" + folder,commonInstallerDestination + "//patches//" + productUpdate + "//" + folder)

    return

# Download caz patch function for SDM(Cumulative only), xFlow, Catalog, ITAM and USS products
def download_Caz(patchDestination,ftpServer,userID,decryptedPassword,PatchSource,cumulativeUpdate,commonInstallerDestination,productUpdate,folderVersion):
    # Download Patch Files
    if not os.path.exists(patchDestination):
        os.makedirs(patchDestination)
        os.chdir(patchDestination)

    ftp = FTP(ftpServer)
    ftp.login(userID, decryptedPassword)

    ftp.cwd(PatchSource)

    file = open(cumulativeUpdate, 'wb')
    logging.info("Attempting to download " + productUpdate + " Cumulative Update to " + patchDestination + "//" + cumulativeUpdate)
    ftp.retrbinary('RETR ' + cumulativeUpdate, file.write)
    file.close()
    ftp.close()
    logging.info(productUpdate + " Cumulative patch file download complete")

    # Copy patch caz files to Common Installer folders.
    if("SDM" not in productUpdate) and ("xFlow" not in productUpdate):
        os.chdir(patchDestination)
        #currentFolderVersion = str(os.listdir(commonInstallerDestination + "//patches//" + productUpdate + "//")[0])
        currentFolderVersion = os.listdir(commonInstallerDestination + "//patches//" + productUpdate + "//")
        for version in currentFolderVersion:
            if (version != folderVersion):
                shutil.rmtree(commonInstallerDestination + "//patches//" + productUpdate + "//" + version)
                logging.info("Current " + productUpdate + " Folder Version is " + version)
                #logging.info("Copying " + productUpdate + " Folder version from " + version + " to " + folderVersion)
                #shutil.move(commonInstallerDestination + "//patches//" + productUpdate + "//" + currentFolderVersion,commonInstallerDestination + "//patches//" + productUpdate + "//" + folderVersion)
        if not os.path.exists(commonInstallerDestination + "//patches//" + productUpdate + "//" + folderVersion + "//Binaries"):
            os.makedirs(commonInstallerDestination + "//patches//" + productUpdate + "//" + folderVersion + "//Binaries")
        shutil.copy(cumulativeUpdate,commonInstallerDestination + "//patches//" + productUpdate + "//" + folderVersion + "//Binaries")
        logging.info(productUpdate + " Patch copy complete....")

    return

def load_properties(filePath, sep='=', comment_char='#'):
    # Read the file passed as parameter as a properties file.

    props = {}
    with open(filePath, "rt") as f:
        for line in f:
            l = line.strip()
            if l and not l.startswith(comment_char):
                key_value = l.split(sep)
                key = key_value[0].strip()
                value = sep.join(key_value[1:]).strip().strip('"')
                props[key] = value
    return props


def updateResponseFiles(filePath, oldLineToReplace, newUpdatedLine):
    with fileinput.FileInput(filePath, inplace=True) as file:
        for line in file:
            print(line.replace(oldLineToReplace, newUpdatedLine), end='')
    return

# Update SDM Patch XML
def updateSDMPatchXML(commonInstallerDestination, patchVersion, majorVersion, minorVersion, rollupVersion,sdmCumulativeUpdate, sdmLocaleUpdate, folderVersion):

    logging.info("Removing existing SDM_Patch.xml and copying from local folder")
    os.remove(commonInstallerDestination + "//patches//SDM_patch.xml")
    shutil.copy("c:\SDM_Patch\SDM_patch.xml", commonInstallerDestination + "//patches//SDM_patch.xml")


    tree = ET.parse(commonInstallerDestination + '//patches//SDM_patch.xml')
    root = tree.getroot()

    for child in root:
        # print(child.attrib)
        if (patchVersion in child.get('id')):
            for elem in root.iter('patches'):
                elem.set('latest', patchVersion)

            for elem in child.iter('patch'):
                elem.set('id', patchVersion)
                elem.set('text', 'Service Desk Manager ' + patchVersion)

            for elem in child.iter('versionInfo'):
                elem.set('major', majorVersion)
                elem.set('minor', minorVersion)
                elem.set('rollup', rollupVersion)

            for elem in child.iter('patchInfo'):
                elem.set('nodeText', 'Service Desk Manager ' + patchVersion)

            for elem in child.iter('binaryPatch'):
                elem.set('fileName', sdmCumulativeUpdate[:-4])

            for elem in child.iter('localePatch'):
                elem.set('fileName', sdmLocaleUpdate[:-4])

            for elem in child.iter('path'):
                elem.set('location', 'patches\\SDM\\' + folderVersion)

            for elem in child.iter('language'):

                if ("en-US" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_en_US")
                elif ("pt-BR" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_pt-BR")
                elif ("fr-CA" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_fr-CA")
                elif ("fr-FR" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_fr-FR")
                elif ("de-DE" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_de-DE")
                elif ("it-IT" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_it-IT")
                elif ("ja-JP" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_ja-JP")
                elif ("es-ES" in elem.get("code")):
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_es-ES")
                else:
                    # print(elem.getAttribute("code"))
                    elem.set("patch", sdmCumulativeUpdate[:-10] + "_zh-CN")

    tree.write(commonInstallerDestination + '//patches//SDM_patch.xml')
    logging.info("Completed updating SDM_patch.xml")
    return

# Update Patch XML Files except SDM

def updateComponentPatchXML(commonInstallerDestination, tempStorageLocation , xmlFilePath, patchVersion, folderVersion, PatchDestination,majorVersion,minorVersion,rollupVersion):

    if ('COLLABSRVR_patch' in xmlFilePath):
        componentDir = "CollaborationServer"
        productInXML = "Collaboration Server"
        PatchDestination = PatchDestination + "//xFlow//"  + componentDir + "//Binaries//"
        xmlLocationUpdate = 'patches\\xFlow\\'
    elif ('XFLOW_patch' in xmlFilePath):
        componentDir = "xFlowAnalyst"
        productInXML = "Xflow"
        PatchDestination = PatchDestination + "//xFlow//"  + componentDir + "//Binaries//"
        xmlLocationUpdate = 'patches\\xFlow\\'
    elif ('SEARCHSRVR_patch' in xmlFilePath):
        componentDir = "SearchServer"
        productInXML = "Search Server"
        PatchDestination = PatchDestination + "//xFlow//"  + componentDir + "//Binaries//"
        xmlLocationUpdate = 'patches\\xFlow\\'
    elif ('USS_patch' in xmlFilePath):
        componentDir = "SelfService"
        productInXML = "Unified Self Service"
        xmlLocationUpdate = 'patches\\'
    elif ('SLCM_patch' in xmlFilePath):
        componentDir = "SLCM"
        productInXML = "SLCM"
        xmlLocationUpdate = 'patches\\'
    else:
        componentDir = "ITAM"
        productInXML = "Asset Portfolio Management"
        xmlLocationUpdate = 'patches\\'

    if(componentDir == "xFlowAnalyst"):
        logging.info("Removing existing " + commonInstallerDestination + "//patches//XFLOW_patch.xml" + " and copying from c:\SDM_Patch\XFLOW_patch.xml")
        logging.info(tempStorageLocation + "//XFLOW_patch.xml")
        logging.info(commonInstallerDestination + "//patches//XFLOW_patch.xml")
        if(os.path.exists(commonInstallerDestination + "//patches//XFLOW_patch.xml")):
            os.remove(commonInstallerDestination + "//patches//XFLOW_patch.xml")
        logging.info("Copying " + tempStorageLocation + "//XFLOW_patch.xml to " + commonInstallerDestination + "//patches//XFLOW_patch.xml")
        shutil.copy(tempStorageLocation + "//XFLOW_patch.xml",commonInstallerDestination + "//patches//XFLOW_patch.xml")

    if (componentDir == "CollaborationServer"):
        logging.info("Removing existing " + commonInstallerDestination + "//patches//COLLABSRVR_patch.xml" + " and copying from c:\SDM_Patch\COLLABSRVR_patch.xml")
        logging.info(tempStorageLocation + "//COLLABSRVR_patch.xml")
        logging.info(commonInstallerDestination + "//patches//COLLABSRVR_patch.xml")
        if(os.path.exists(commonInstallerDestination + "//patches//COLLABSRVR_patch.xml")):
            os.remove(commonInstallerDestination + "//patches//COLLABSRVR_patch.xml")
        logging.info("Copying " + tempStorageLocation + "//COLLABSRVR_patch.xml to " + commonInstallerDestination + "//patches//COLLABSRVR_patch.xml")
        shutil.copy(tempStorageLocation + "//COLLABSRVR_patch.xml",commonInstallerDestination + "//patches//COLLABSRVR_patch.xml")

    if (componentDir == "SearchServer"):
        logging.info("Removing existing " + commonInstallerDestination + "//patches//SEARCHSRVR_patch.xml" + " and copying from c:\SDM_Patch\SEARCHSRVR_patch.xml")
        logging.info(tempStorageLocation + "//SEARCHSRVR_patch.xml")
        logging.info(commonInstallerDestination + "//patches//SEARCHSRVR_patch.xml")
        if(os.path.exists(commonInstallerDestination + "//patches//SEARCHSRVR_patch.xml")):
            os.remove(commonInstallerDestination + "//patches//SEARCHSRVR_patch.xml")
        logging.info("Copying " + tempStorageLocation + "//SEARCHSRVR_patch.xml to " + commonInstallerDestination + "//patches//SEARCHSRVR_patch.xml")
        shutil.copy(tempStorageLocation + "//SEARCHSRVR_patch.xml",commonInstallerDestination + "//patches//SEARCHSRVR_patch.xml")

    if (componentDir == "SLCM"):
        logging.info("Removing existing " + commonInstallerDestination + "//patches//SLCM_patch.xml" + " and copying from c:\SDM_Patch\SLCM_patch.xml")
        logging.info(tempStorageLocation + "//SLCM_patch.xml")
        logging.info(commonInstallerDestination + "//patches//SLCM_patch.xml")
        if(os.path.exists(commonInstallerDestination + "//patches//SLCM_patch.xml")):
            os.remove(commonInstallerDestination + "//patches//SLCM_patch.xml")
        logging.info("Copying " + tempStorageLocation + "//SLCM_patch.xml to " + commonInstallerDestination + "//patches//SLCM_patch.xml")
        shutil.copy(tempStorageLocation + "//SLCM_patch.xml", commonInstallerDestination + "//patches//SLCM_patch.xml")

    if (componentDir == "ITAM"):
        logging.info("Removing existing " + commonInstallerDestination + "//patches//ITAM_patch.xml" +  " and copying from " + "c://SDM_Patch//ITAM_patch.xml")
        if(os.path.exists(commonInstallerDestination + "//patches//ITAM_patch.xml")):
            os.remove(commonInstallerDestination + "//patches//ITAM_patch.xml")
        logging.info("Copying " + tempStorageLocation + "//ITAM_patch.xml to " + commonInstallerDestination + "//patches//ITAM_patch.xml")
        shutil.copy(tempStorageLocation + "//ITAM_patch.xml", commonInstallerDestination + "//patches//ITAM_patch.xml")

    if (componentDir == "SelfService"):
        logging.info("Removing existing " + commonInstallerDestination + "//patches//USS_patch.xml" + " and copying from c://SDM_Patch//USS_patch.xml")
        logging.info(tempStorageLocation + "//USS_patch.xml")
        logging.info(commonInstallerDestination + "//patches//USS_patch.xml")
        if(os.path.exists(commonInstallerDestination + "//patches//USS_patch.xml")):
            os.remove(commonInstallerDestination + "//patches//USS_patch.xml")
        logging.info("Copying " + tempStorageLocation + "//USS_patch.xml to " + commonInstallerDestination + "//patches//USS_patch.xml")
        shutil.copy(tempStorageLocation + "//USS_patch.xml",commonInstallerDestination + "//patches//USS_patch.xml")

    tree = ET.parse(xmlFilePath)
    root = tree.getroot()
    for child in root:
        # print(child.attrib)
        if (patchVersion in child.get('id')):
            boolValue = 0
            for elem in root.iter('patches'):
                elem.set('latest', patchVersion)

            for elem in child.iter('patch'):
                elem.set('id', patchVersion)

            for elem in child.iter('patchInfo'):
                elem.set('nodeText', productInXML + " " + patchVersion)

            for elem in child.iter('versionInfo'):
                elem.set('major', majorVersion)
                elem.set('minor', minorVersion)
                elem.set('rollup', rollupVersion)

            component = os.listdir(PatchDestination)
            for patchName in component:
                logging.info("Component Patch binary is " + patchName)
                for elem in child.iter('binaryPatch'):
                    elem.set("fileName", patchName[:-4])

                for elem in child.iter('path'):
                    elem.set('location', xmlLocationUpdate + componentDir + '\\' + folderVersion)

    tree.write(xmlFilePath)
    return

#VARIABLES USED FOR YAML : ftpServer, userID, decryptedPassword, sdmProductUpdate, xflowProductUpdate, commonInstallerSource, commonInstallerFile, commonInstallerDestination, sdmPatchSource,
# xflowPatchSource, sdmCumulativeUpdate, xflowCumulativeUpdate, sdmPatchDestination, sdmLocalePatchDestination, xflowPatchDestination, majorVersion, minorVersion, patchVersion, folderVersion
# sdmXmlFile, xflowXmlFile, responseFilePath

def yaml_loader(filepath):
    #Loading yaml file
    with open(filepath,"r") as file_descriptor:
        data = yaml.load(file_descriptor)
        return data

if __name__ == "__main__":
    if os.name =="nt":
      os.chdir("C://SDM_Patch//")
    else:
      os.chdir("//SDM_Patch//")

    filepath = "config.yaml"
    logging.info("Attempting to load config.yaml")
    logging.info("-------------------------------")
    logging.info("Reading YAML entries.......")
    data = yaml_loader(filepath)


items = data.get('SDM_Autodeploy')
for itemName, itemValue in items.items():
    if ("FTP_Server" in itemName):
        ftpServer = itemValue
        logging.info("ftpServer is " + ftpServer)

items = data.get('Domain_Credentials')

for itemName,itemValue in items.items():
     if("User_ID" in itemName):
         userID = itemValue
         logging.info("Domain User ID is " + itemValue)
     if("Password" in itemName):
             decryptedPassword = base64.b64decode(itemValue).decode('utf-8')

items = data.get('Products_to_update')
for itemName,itemValue in items.items():
     if("SDM" in itemName):
         sdmProductUpdate = itemValue
         if(sdmProductUpdate=="SDM"):
            logging.info("SDM is to be updated")
     if("xFlow" in itemName):
            xflowProductUpdate = itemValue
            if (xflowProductUpdate == "xFlow"):
                logging.info("xFlow is to be updated")
     if ("SelfService" in itemName):
            ussProductUpdate = itemValue
            if (ussProductUpdate == "SelfService"):
                logging.info("USS is to be updated")
     if ("SLCM" in itemName):
            slcmProductUpdate = itemValue
            if(slcmProductUpdate=="SLCM"):
                logging.info("SLCM is to be updated")
     if ("ITAM" in itemName):
            itamProductUpdate = itemValue
            if(itamProductUpdate=="ITAM"):
                logging.info("ITAM is to be updated")


items = data.get('Patch_Version')
for itemName, itemValue in items.items():
    if ("Major" in itemName):
        majorVersion = str(itemValue)
        logging.info("majorVersion is  " + majorVersion)
    if ("Minor" in itemName):
        minorVersion = str(itemValue)
    if ("Rollup" in itemName):
        rollupVersion = str(itemValue)
        if(rollupVersion=="0"):
           patchVersion = majorVersion + "." + minorVersion
           folderVersion = majorVersion + "_" + minorVersion
        else:
            patchVersion = majorVersion + "." + minorVersion + "." + rollupVersion
            folderVersion = majorVersion + "_" + minorVersion + "_" + rollupVersion
        logging.info("minorVersion is " + minorVersion)
        logging.info("patchVersion is " + patchVersion)
        logging.info("folderVersion is " + folderVersion)

    items = data.get('Update_Mode')
for itemName, itemValue in items.items():
    if ("CI" in itemName):
        ciUpdateMode = itemValue
        logging.info("ciUpdateMode is  " + ciUpdateMode)
    if ("SDM" in itemName):
        sdmUpdateMode = itemValue
        if (sdmProductUpdate == "SDM"):
            logging.info("sdmUpdateMode is  " + sdmUpdateMode)
    if ("Prefix" in itemName):
        sdmPrefix = itemValue
        if (sdmProductUpdate == "SDM"):
            logging.info("sdmPrefix is  " + sdmPrefix)
    if ("xFlow" in itemName):
        xflowUpdateMode = itemValue
        if (xflowProductUpdate == "xFlow"):
            logging.info("xflowUpdateMode is  " + xflowUpdateMode)
    if ("SelfService" in itemName):
        ussUpdateMode = itemValue
        if (ussProductUpdate == "SelfService"):
            logging.info("ussUpdateMode is  " + ussUpdateMode)
    if ("SLCM" in itemName):
        slcmUpdateMode = itemValue
        if (slcmProductUpdate == "SLCM"):
            logging.info("slcmUpdateMode is  " + slcmUpdateMode)
    if ("ITAM" in itemName):
        itamUpdateMode = itemValue
        if (itamProductUpdate == "ITAM"):
            logging.info("itamUpdateMode is  " + itamUpdateMode)

items = data.get('Update_Mechanism')

for itemName,itemValue in items.items():
     if("Type" in itemName):
         updateMechanism = itemValue
         logging.info("Update Mechanism is " + updateMechanism)
     if ("Workspace_Location" in itemName):
         workspaceDir = itemValue
         logging.info("Workspace directory is  " + workspaceDir)

items = data.get('Common_Installer_Source_Path')
for itemName, itemValue in items.items():
    if ("CI_Source" in itemName):
        commonInstallerSource = itemValue
        logging.info("CISource is " + commonInstallerSource)
    if(ciUpdateMode ==  "Auto"):
        # Retrieving files
        host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)

        # Retrieving Common Installer patch file name
        host.chdir(commonInstallerSource)
        folderList = host.listdir(host.curdir)
        commonInstallerFile = folderList[-1]
    else:
     if ("CI_File" in itemName):
        commonInstallerFile = itemValue
        logger.info("CI File is " + commonInstallerFile)

    if ("CI_Destination" in itemName):
        commonInstallerDestination = itemValue
        logging.info("CI_Destination is " + commonInstallerDestination)

items = data.get('Patches_Source_Path')
for itemName, itemValue in items.items():
    if(sdmUpdateMode == "Auto") and (sdmProductUpdate == "SDM"):
        host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
        host.chdir('/SDMBUILDS/SD/NT/')
        #####
        folderList = [d for d in host.listdir('.') if host.path.isdir(d) if sdmPrefix in d]
        latestSubdir = max(folderList, key=host.path.getmtime)
        sdmCumulativeUpdate = latestSubdir + "_cum_C.caz"
        sdmcumulativePatchPrefix = latestSubdir
        sdmPatchSource = '/SDMBUILDS/SD/NT/' + sdmcumulativePatchPrefix
        #####
        #folderList = host.listdir(host.curdir)
        #folderArray = []
        #for folder in folderList:
            #if (sdmPrefix in folder):
                #folderArray.append(folder)
        #print(folderArray[-1])
        #sdmCumulativeUpdate = folderArray[-1] + "_cum_C.caz"
        #sdmcumulativePatchPrefix = folderArray[-1]
        #sdmPatchSource = '/SDMBUILDS/SD/NT/' + sdmcumulativePatchPrefix
    else:
        if ("SDM" in itemName) and (sdmProductUpdate=="SDM"):
            sdmPatchSource = itemValue

    if ("xFlow" in itemName) and (xflowProductUpdate=="xFlow"):
        xflowPatchSource = itemValue
        logging.info("xFlow Patch Source is  " + xflowPatchSource)

    if ("USS" in itemName) and (ussProductUpdate=="SelfService"):
        ussPatchSource = itemValue
        if(ussUpdateMode=="Auto"):
            host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
            host.chdir(ussPatchSource)
            """folderList = host.listdir(host.curdir)
            folderArray = []
            for folder in folderList:
                if ("Build" in folder):
                    folderArray.append(folder)
            ussPatchSource = ussPatchSource + "/" + folderList[-1]"""

            #####
            folderList = [d for d in host.listdir('.') if host.path.isdir(d) if "Build" in d]
            latestSubdir = max(folderList, key=host.path.getmtime)
            ussPatchSource = ussPatchSource + "/" + latestSubdir
            ######

    if ("SLCM" in itemName) and (slcmProductUpdate == "SLCM"):
        ###
        slcmPatchSource = itemValue
        if(slcmUpdateMode=="Auto"):
            host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
            host.chdir(slcmPatchSource)
            """folderList = host.listdir(host.curdir)
            folderArray = []
            for folder in folderList:
                if ("Build" in folder):
                    folderArray.append(folder)
            slcmPatchSource = slcmPatchSource + "/" + folderList[-1] """
            #####
            folderList = [d for d in host.listdir('.') if host.path.isdir(d) if "Build" in d]
            latestSubdir = max(folderList, key=host.path.getmtime)
            slcmPatchSource = slcmPatchSource + "/" + latestSubdir
            ######



    if ("ITAM" in itemName) and (itamProductUpdate=="ITAM"):
        itamPatchSource = itemValue
        if(itamUpdateMode=="Auto"):
            host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
            host.chdir(itamPatchSource)
            """folderList = host.listdir(host.curdir)
            folderArray = []
            for folder in folderList:
                if ("Build" in folder):
                    folderArray.append(folder)
            itamPatchSource = itamPatchSource + "/" + folderList[-1]"""
            #####
            folderList = [d for d in host.listdir('.') if host.path.isdir(d) if "Build" in d]
            latestSubdir = max(folderList, key=host.path.getmtime)
            itamPatchSource = itamPatchSource + "/" + latestSubdir
            ######

if(sdmProductUpdate == "SDM"):
    logging.info("SDM Patch Source is " + sdmPatchSource)

if(slcmProductUpdate == "SLCM"):
    logging.info("SLCM Patch Source is " + slcmPatchSource)

if(itamProductUpdate == "ITAM"):
    logging.info("ITAM Patch Source is  " + itamPatchSource)

if (ussProductUpdate == "SelfService"):
    logging.info("USS Patch Source is  " + ussPatchSource)

items = data.get('Cumulative_Update_File')
for itemName, itemValue in items.items():
    if(sdmUpdateMode != "Auto") and (sdmProductUpdate == "SDM"):
      if ("SDM_Cum" in itemName) and (sdmProductUpdate=="SDM"):
          sdmCumulativeUpdate = itemValue
          logging.info("SDM Cumulative Update is " + sdmCumulativeUpdate)
    if ("SDM_Locale" in itemName) and (sdmProductUpdate=="SDM"):
        sdmLocaleUpdate = itemValue
        logging.info("SDM Locale Update is " + sdmLocaleUpdate)
    if(xflowUpdateMode=="Auto")  and (xflowProductUpdate=="xFlow"):
        host.chdir(xflowPatchSource)
        folderList = host.listdir(host.curdir)
        xflowCumulativeUpdate = folderList[-1]
    else:
         if("xFlow" in itemName) and (xflowProductUpdate=="xFlow"):
          xflowCumulativeUpdate = itemValue
    if ("USS" in itemName) and (ussProductUpdate == "SelfService"):
        if (ussUpdateMode == "Auto"):
            host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
            host.chdir(ussPatchSource)
            folderList = host.listdir(host.curdir)
            ussCumulativeUpdate = folderList[-1]
        else:
            ussCumulativeUpdate = itemValue
        logging.info("USS Cumulative Update is " + ussCumulativeUpdate)
    if ("SLCM" in itemName) and (slcmProductUpdate == "SLCM"):
        if (slcmUpdateMode == "Auto"):
            host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
            host.chdir(slcmPatchSource)
            folderList = host.listdir(host.curdir)
            folderArray = []
            for folder in folderList:
                if (".CAZ" in folder):
                    folderArray.append(folder)
                    slcmCumulativeUpdate = folderArray[-1]
        else:
            slcmCumulativeUpdate = itemValue
        logging.info("SLCM Cumulative Update is " + slcmCumulativeUpdate)
    if ("ITAM" in itemName) and (itamProductUpdate == "ITAM"):
        if (itamUpdateMode == "Auto"):
            host = ftputil.FTPHost(ftpServer, userID, decryptedPassword)
            host.chdir(itamPatchSource)
            folderList = host.listdir(host.curdir)
            itamCumulativeUpdate = folderList[-1]
        else:
            itamCumulativeUpdate = itemValue
        logging.info("ITAM Cumulative Update is " + itamCumulativeUpdate)

if(sdmProductUpdate=="SDM"):
    logging.info("SDM Cumulative Update is " + sdmCumulativeUpdate)
if(xflowProductUpdate=="xFlow"):
    logging.info("xFlow Cumulative Update is " + xflowCumulativeUpdate)

items = data.get('Patches_Destination_Path')
for itemName, itemValue in items.items():
    if ("SDM_Cum" in itemName)  and (sdmProductUpdate == "SDM"):
        sdmPatchDestination = itemValue
        logging.info("SDM Patch Destination is " + sdmPatchDestination)
    if ("SDM_Locale" in itemName)  and (sdmProductUpdate == "SDM"):
        sdmLocalePatchDestination = itemValue
        logging.info("SDM Locale Patch Destination is " + sdmLocalePatchDestination)
    if ("xFlow" in itemName)  and (xflowProductUpdate == "xFlow"):
        xflowPatchDestination = itemValue
        logging.info("xFlow Patch Destination is  " + xflowPatchDestination)
    if ("USS" in itemName)  and (ussProductUpdate == "SelfService"):
        ussPatchDestination = itemValue
        logging.info("USS Patch Destination is  " + ussPatchDestination)
    if ("SLCM" in itemName)  and (slcmProductUpdate == "SLCM"):
        slcmPatchDestination = itemValue
        logging.info("SLCM Patch Destination is  " + slcmPatchDestination)
    if ("ITAM" in itemName)  and (itamProductUpdate == "ITAM"):
        itamPatchDestination = itemValue
        logging.info("ITAM Patch Destination is  " + itamPatchDestination)

items = data.get('XMLFiles_to_update')
for itemName, itemValue in items.items():
    if ("SDM" in itemName)  and (sdmProductUpdate == "SDM"):
        sdmXmlFile = itemValue
        logging.info("SDM XML File is " + sdmXmlFile)
    if ("xFlow" in itemName)  and (xflowProductUpdate == "xFlow"):
        xflowXmlFile = itemValue
        logging.info("xFlow XML File is  " + xflowXmlFile)
    if ("collabSrvr" in itemName)  and (xflowProductUpdate == "xFlow"):
        collabSrvrXmlFile = itemValue
        logging.info("collabSrvr XML File is  " + collabSrvrXmlFile)
    if ("searchSrvr" in itemName)  and (xflowProductUpdate == "xFlow"):
        searchSrvrXmlFile = itemValue
        logging.info("searchSrvr XML File is  " + searchSrvrXmlFile)
    if ("USS" in itemName)  and (ussProductUpdate == "SelfService"):
        ussXmlFile = itemValue
        logging.info("USS XML File is  " + ussXmlFile)
    if ("SLCM" in itemName)  and (slcmProductUpdate == "SLCM"):
        slcmXmlFile = itemValue
        logging.info("SLCM XML File is  " + slcmXmlFile)
    if ("ITAM" in itemName)  and (itamProductUpdate == "ITAM"):
        itamXmlFile = itemValue
        logging.info("ITAM XML File is  " + itamXmlFile)

items = data.get('Installation_Type')
for itemName, itemValue in items.items():
    if ("Fresh_Install" in itemName):
        freshInstall = itemValue
        logging.info("Fresh Install is set to " + freshInstall)
    if ("Install_Source" in itemName):
        productInstallSource = itemValue
        logging.info("productInstallSource  is " + productInstallSource)

items = data.get('Setup_Response_Files')
for itemName, itemValue in items.items():
    if ("Response_Path" in itemName):
        responseFilePath = itemValue
        logging.info("Response Files Path is " + responseFilePath)
logging.info("Completed reading YAML settings")

#Open ftp connection

ftp = FTP(ftpServer)
ftp.login(userID,decryptedPassword)


# Download Common Installer

"""if os.name=="nt":
   if os.path.exists('c://Workspace'):
      shutil.rmtree('c://Workspace')
else:
    if os.path.exists('//Workspace'):
        shutil.rmtree('//Workspace')"""

if os.path.exists(workspaceDir):
      shutil.rmtree(workspaceDir)

if not os.path.exists(commonInstallerDestination):
    os.makedirs(commonInstallerDestination)
    os.chdir(commonInstallerDestination)

ftp.cwd(commonInstallerSource)

file = open(commonInstallerFile, 'wb')
logging.info("Attempting to download Common Installer file to " + commonInstallerDestination + "//" + commonInstallerFile)
ciFileSize = ftp.size(commonInstallerFile)
ftp.retrbinary('RETR '+ commonInstallerFile, file.write)
ftp.close()
file.close()

if(sdmProductUpdate=="SDM"):
        # Download Patch Files
        # SDM Cumulative
        download_Caz(sdmPatchDestination, ftpServer, userID, decryptedPassword, sdmPatchSource, sdmCumulativeUpdate, commonInstallerDestination, sdmProductUpdate, folderVersion)

        # SDM Locale
        ftp = FTP(ftpServer)
        ftp.login(userID, decryptedPassword)
        ftp.cwd(sdmPatchSource)
        if not os.path.exists(sdmLocalePatchDestination):
            os.makedirs(sdmLocalePatchDestination)
            os.chdir(sdmLocalePatchDestination)

        filematch ='*.caz'

        # Loop through matching files and download each one individually
        for filename in ftp.nlst(filematch):
            if "cum" not in filename and "TEST" not in filename and "COMBO" not in filename and "ROLLUP" not in filename:
                fhandle = open(filename, 'wb')
                logging.info('Downloading locale patch ' + filename)
                ftp.retrbinary('RETR ' + filename, fhandle.write)
                fhandle.close()
                logging.info(filename + " download complete")

if(xflowProductUpdate=="xFlow"):
    # xFlow
    download_Caz(xflowPatchDestination, ftpServer, userID, decryptedPassword, xflowPatchSource, xflowCumulativeUpdate,commonInstallerDestination, xflowProductUpdate, folderVersion)

# Extracting Common Installer and xFlow zip
if os.name =="nt":
    if (ciFileSize == os.path.getsize(commonInstallerDestination + "//" + commonInstallerFile)):
       logging.info("Downloaded Common Installer Successfully")
       logging.info("Attempting to unzip the Common Installer...")
       zip_ref = zipfile.ZipFile(commonInstallerDestination + "//" + commonInstallerFile, 'r')
       zip_ref.extractall(commonInstallerDestination)
       zip_ref.close()
       logging.info("Common Installer unzipped successfully")
       if(xflowProductUpdate=="xFlow"):
            logging.info("Attempting to unzip xFlow Product Update")
            zip_ref = zipfile.ZipFile(xflowPatchDestination + "//" + xflowCumulativeUpdate, 'r')
            zip_ref.extractall(xflowPatchDestination)
            zip_ref.close()
            logging.info("xFlow Product Update unzipped successfully")

if("17.1" not in commonInstallerSource):
    commonInstallerDestination = commonInstallerDestination + "//CASM_DVD"
    logging.info("Updated commonInstaller Destination is " + commonInstallerDestination)

if(sdmProductUpdate=="SDM"):
    createFolderStructure(tempStorageLocation, folderVersion, sdmProductUpdate, commonInstallerDestination)

if(xflowProductUpdate=="xFlow"):
    createFolderStructure(tempStorageLocation + "//xFlow//xFlowAnalyst//", folderVersion, xflowProductUpdate + "//xFlowAnalyst//", commonInstallerDestination)
    createFolderStructure(tempStorageLocation + "//xFlow//CollaborationServer//", folderVersion, xflowProductUpdate + "//CollaborationServer//", commonInstallerDestination)
    createFolderStructure(tempStorageLocation + "//xFlow//SearchServer//", folderVersion, xflowProductUpdate + "//SearchServer//", commonInstallerDestination)

    listOfXflowProducts = os.listdir(commonInstallerDestination + "//patches//" + xflowProductUpdate + "//")
    for product in listOfXflowProducts:
        if(os.path.exists(commonInstallerDestination + "//patches//" + xflowProductUpdate + "//" + product + "//" + folderVersion + "//Binaries")):
            shutil.rmtree(commonInstallerDestination + "//patches//" + xflowProductUpdate + "//" + product + "//" + folderVersion + "//Binaries")
        if os.path.exists(xflowPatchDestination + "//xFlow//" + product + "//Binaries"):
            shutil.copytree(xflowPatchDestination + "//xFlow//" + product + "//Binaries", commonInstallerDestination + "//patches//" + xflowProductUpdate + "//" + product + "//" + folderVersion + "//Binaries")
            logging.info("Completed copying " + product + " Folder to " + folderVersion)
        else:
            shutil.rmtree(commonInstallerDestination + "//patches//" + xflowProductUpdate + "//" + product)

if(sdmProductUpdate=="SDM"):
    # Download SDM Locale caz
    """if os.name=="nt":
        logging.info("Attempting to download Cazipxp.exe to create SDM Locale binary")
        os.chdir(sdmLocalePatchDestination)

        ftp = FTP("ftp.ca.com")
        ftp.login()
        ftp.cwd("/CAproducts")

        file = open("Cazipxp.exe", 'wb')
        logging.info("Attempting to download Cazipxp.exe to " + sdmLocalePatchDestination)
        ftp.retrbinary('RETR Cazipxp.exe', file.write)
        file.close()
        ftp.quit()
        logging.info("Completed download cazipxp.exe successfully")"""

    #Create SDM Locale Caz
if(updateMechanism!="ApplyPTF") and (sdmProductUpdate=="SDM"):
    logging.info("Creating SDM Language pack now....")

    sdmLocalePatchDestination = sdmLocalePatchDestination
    f = open(sdmLocalePatchDestination + '//' + majorVersion + minorVersion + 'Testing.JCL', 'w+')
    f.write('PLATFORM:WINDOWS\n')
    f.write('PRODUCT:USRD\n')
    f.write('COMPONENT:USRD-CMN\n')
    f.write('SUPERSEDE:\n')
    f.write('RELEASE:17.1\n')
    f.write('GENLEVEL:0000\n')
    f.write('VERSION:20010222\n')
    f.write('PREREQS:\n')
    f.write('MPREREQS:\n')
    f.write('COREQS:\n')
    f.write('MCOREQS:\n')
    f.write('\n')
    f.write('\n')

    folderList = os.listdir(sdmLocalePatchDestination)
    for fileName in folderList:
       if ('.caz' in fileName) and ('Test' not in fileName):
          logging.info("Adding language pack : " + fileName)
          f.write('FILE:patches\\' + fileName + ':NEWFILE:\n')
    f.close()

    if os.path.exists(sdmLocalePatchDestination + '//' + majorVersion + minorVersion + 'Testing.caz'):
       os.remove(sdmLocalePatchDestination + '//' + majorVersion + minorVersion + 'Testing.caz')

    os.chdir(sdmLocalePatchDestination)
    command = commonInstallerDestination + "//filestore//utils//CAZIP//cazipxp -w *.caz " + majorVersion + minorVersion + 'Testing.caz'
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("New Caz File created with exit code " + str(p.returncode))

    command = commonInstallerDestination + "//filestore//utils//CAZIP//cazipxp -a *.JCL " + majorVersion + minorVersion + 'Testing.caz'
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("Caz File updated with JCL file and returned exit code " + str(p.returncode))

    # Copy patch caz files to Common Installer folders.
    os.chdir(sdmPatchDestination)
    if not os.path.exists(commonInstallerDestination + "//patches//SDM//" + folderVersion + "//Binaries"):
        os.makedirs(commonInstallerDestination + "//patches//SDM//" + folderVersion + "//Binaries")
    shutil.copy(sdmCumulativeUpdate,commonInstallerDestination + "//patches//SDM//" + folderVersion + "//Binaries")
    os.chdir(sdmLocalePatchDestination)
    if not os.path.exists(commonInstallerDestination + "//patches//SDM//" + folderVersion + "//Locale"):
        os.makedirs(commonInstallerDestination + "//patches//SDM//" + folderVersion + "//Locale")
    shutil.copy(sdmLocaleUpdate,commonInstallerDestination + "//patches//SDM//" + folderVersion + "//Locale")
    logging.info("SDM Language pack created successfully....")

# Download USS Caz
if(ussProductUpdate=="SelfService"):
    download_Caz(ussPatchDestination,ftpServer, userID, decryptedPassword, ussPatchSource, ussCumulativeUpdate, commonInstallerDestination, ussProductUpdate, folderVersion)

# Download SLCM Caz
if(slcmProductUpdate=="SLCM"):
    download_Caz(slcmPatchDestination, ftpServer, userID, decryptedPassword, slcmPatchSource, slcmCumulativeUpdate,commonInstallerDestination, slcmProductUpdate, folderVersion)

# Download ITAM Caz
if(itamProductUpdate=="ITAM"):
    download_Caz(itamPatchDestination, ftpServer, userID, decryptedPassword, itamPatchSource, itamCumulativeUpdate,commonInstallerDestination, itamProductUpdate, folderVersion)

# Start ApplyPTF based update
if(updateMechanism=="ApplyPTF"):
    logging.info("Found Update Mechanism is through AppyPTF - not attempting to update Common Installer XML files and Response Properties")
    mdbUpdate = False
    sdmPatchDestination = sdmPatchDestination.replace("/", "\\")
    sdmMdbPatchDestination =  sdmPatchDestination + "\\MDB"
    sdmLocalePatchDestination = sdmLocalePatchDestination.replace("/", "\\")
    commonInstallerDestination = commonInstallerDestination.replace("/", "\\")
    patchProps = load_properties(os.environ['systemroot'] + "//paradigm.ini", "=", "#")
    sdmInstalledLocation = patchProps['NX_ROOT']
    sdmInstalledLocation = sdmInstalledLocation.replace("/", "\\")
    logging.info("SDM Installed Location is " + sdmInstalledLocation)

    patchProps = load_properties(sdmInstalledLocation + "//NX.env", "=", "#")
    sdmLocale = patchProps['@NX_CA_BOOKSHELF_LANG']
    jreDir = patchProps['@NX_JRE_INSTALL_DIR']
    sdmDBType = patchProps['@NX_DB_TYPE']
    sdmDBName = patchProps['@NX_DB_STUFF']
    sdmDBHost = patchProps['@NX_DB_NODE']
    sdmDBPort = patchProps['@NX_DB_PORT']

    logging.info("SDM Product locale is " + sdmLocale)
    logging.info("SDM DB is " + sdmDBType)
    logging.info("SDM DBHost is " + sdmDBHost)
    logging.info("SDM DBPort is " + sdmDBPort)
    logging.info("JRE Directory is " + jreDir)

    # Extracting SDM Cumulative
    patchExtraction(sdmPatchDestination, commonInstallerDestination, sdmCumulativeUpdate)
    time.sleep(20)

    # Determine if patch includes MDB updates and extract updates
    os.chdir(sdmPatchDestination)
    currentFolderList = os.listdir(sdmPatchDestination)
    for eachFile in currentFolderList:
        if ("ORACLE_MDB" in eachFile):
            logging.info("Patch includes an MDB update")
            mdbUpdate = True
            if not os.path.exists(sdmMdbPatchDestination):
                os.makedirs(sdmMdbPatchDestination)
                if (sdmDBType == "SQL"):
                    shutil.copyfile("MSSQL_MDB.CAZ", sdmMdbPatchDestination + "//MSSQL_MDB.CAZ")
                    sdmMdbUpdate = "MSSQL_MDB.CAZ"
                else:
                    shutil.copyfile("ORACLE_MDB.CAZ", sdmMdbPatchDestination + "//ORACLE_MDB.CAZ")
                    sdmMdbUpdate = "ORACLE_MDB.CAZ"
            patchExtraction(sdmMdbPatchDestination, commonInstallerDestination, sdmMdbUpdate)

    if (sdmDBType == "SQL") and (mdbUpdate == True):
        os.chdir(sdmMdbPatchDestination)
        mdbSetupCommand = "setupmdb.bat -DBVENDOR=" + '"{}"'.format("mssql") + " -DBNAME=" + '"{}"'.format(
            sdmDBName) + " -DBHOST=" + '"{}"'.format(sdmDBHost) + " -DBPORT=" + '"{}"'.format(
            sdmDBPort) + " -DBUSER=" + '"{}"'.format("sa") + " -DBPASSWORD=" + '"{}"'.format(
            "N0tallowed") + " -JRE_DIR=" + '"{}"'.format(jreDir) + " -WORKSPACE=" + '"{}"'.format(
            "Service_Desk") + " -DBDRIVER=" + '"{}"'.format("Service_Desk")
        logging.info("mdb setup Command is " + mdbSetupCommand)
        p = subprocess.Popen(mdbSetupCommand, shell=True, stdout=subprocess.PIPE)
        stdout, stderr = p.communicate()
        logging.info("SetupMDB completed with exit code " + str(p.returncode))
        time.sleep(20)
    # Add Oracle step here

    # Extracting SDM Locale
    sdmLocale = sdmLocale.replace('-', '_')
    sdmLocaleUpdate = sdmcumulativePatchPrefix  + "_" + sdmLocale + ".caz"
    logging.info("SDM Locale patch is " + sdmLocaleUpdate)
    patchExtraction(sdmLocalePatchDestination, commonInstallerDestination, sdmLocaleUpdate)
    time.sleep(20)

    # Executing post install steps
    ## Copy required customization files
    copyWebXmlFiles("web.xml", "C://SDM_Patch//web.xml",
                    sdmInstalledLocation + "//bopcfg//www//CATALINA_BASE//webapps//CAisd//WEB-INF//web.xml")
    copyWebXmlFiles("web.xml.tpl", "C://SDM_Patch//web.xml.tpl",
                    sdmInstalledLocation + "//samples//pdmconf//web.xml.tpl")
    time.sleep(20)

    # Start SDM Service
    command = "net start pdm_daemon_manager"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("SDM Service start completed with exit code " + str(p.returncode))
    time.sleep(45)

    # PDM_configure
    command = "pdm_configure - S"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("PDM_configure completed with exit code " + str(p.returncode))
    time.sleep(20)

    # Run PDM_Load
    applyDatFiles(sdmPatchDestination, sdmInstalledLocation, "insert.dat")
    applyDatFiles(sdmPatchDestination, sdmInstalledLocation, "update.dat")
    applyDatFiles(sdmPatchDestination, sdmInstalledLocation, "delete.dat")
    time.sleep(20)

    # PDM_WebCache to alert users to clear their browser cache - Browser refresh time set
    logging.info("Executing the PDM WebCache Command")
    command = "pdm_webcache -b"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("PDM_webcache completed with exit code " + str(p.returncode))
    time.sleep(20)

    # PDM_Options_mgr
    logging.info("Executing the PDM_options_mgr Command")
    command = "pdm_options_mgr"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("PDM_options_mgr completed with exit code " + str(p.returncode))
    time.sleep(20)

    command = "pdm_configure - S"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("PDM_configure completed with exit code " + str(p.returncode))
    time.sleep(20)
    # Stop and Start SDM ser service

    command = "net stop pdm_daemon_manager"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("SDM Service stop completed with exit code " + str(p.returncode))
    time.sleep(20)

    command = "net start pdm_daemon_manager"
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("SDM Service start completed with exit code " + str(p.returncode))
    time.sleep(20)
    sys.exit()

# End ApplyPTF based update

# Update SDM Patch XML File
if(sdmProductUpdate == "SDM"):
    if(os.path.exists(sdmXmlFile)):
        updateSDMPatchXML(commonInstallerDestination,patchVersion,majorVersion,minorVersion,rollupVersion,sdmCumulativeUpdate,sdmLocaleUpdate,folderVersion)

# Update other product XML files
if(xflowProductUpdate=="xFlow"):
    if os.path.exists(xflowPatchDestination + "//xFlow//CollaborationServer//Binaries//"):
        if os.path.exists(collabSrvrXmlFile):
            logging.info("Updating CollabServer Components XML files....")
            updateComponentPatchXML(commonInstallerDestination, tempStorageLocation , collabSrvrXmlFile, patchVersion, folderVersion, xflowPatchDestination,majorVersion,minorVersion,rollupVersion)
if(xflowProductUpdate=="xFlow"):
    if os.path.exists(xflowPatchDestination + "//xFlow//SearchServer//Binaries//"):
        if os.path.exists(searchSrvrXmlFile):
            logging.info("Updating Search Server Components XML files....")
            updateComponentPatchXML(commonInstallerDestination, tempStorageLocation, searchSrvrXmlFile, patchVersion, folderVersion, xflowPatchDestination,majorVersion,minorVersion,rollupVersion)
if(xflowProductUpdate=="xFlow"):
    if os.path.exists(xflowPatchDestination + "//xFlow//xFlowAnalyst//Binaries//"):
        if os.path.exists(xflowXmlFile):
            logging.info("Updating xFlow Components XML files....")
            updateComponentPatchXML(commonInstallerDestination, tempStorageLocation, xflowXmlFile, patchVersion, folderVersion, xflowPatchDestination, majorVersion,minorVersion, rollupVersion)
if(ussProductUpdate=="SelfService"):
    if os.path.exists(ussPatchDestination):
        if os.path.exists(ussXmlFile):
            logging.info("Updating USS Patch XML files....")
            updateComponentPatchXML(commonInstallerDestination, tempStorageLocation, ussXmlFile, patchVersion, folderVersion, ussPatchDestination, majorVersion,minorVersion, rollupVersion)
if(slcmProductUpdate=="SLCM"):
    if os.path.exists(slcmPatchDestination):
        if os.path.exists(slcmXmlFile):
            logging.info("Updating SLCM Patch XML file....")
            updateComponentPatchXML(commonInstallerDestination, tempStorageLocation, slcmXmlFile, patchVersion, folderVersion, slcmPatchDestination, majorVersion,minorVersion, rollupVersion)
if(itamProductUpdate=="ITAM"):
    if os.path.exists(itamPatchDestination):
        if os.path.exists(itamXmlFile):
            logging.info("Updating ITAM Patch XML file....")
            updateComponentPatchXML(commonInstallerDestination, tempStorageLocation, itamXmlFile, patchVersion, folderVersion, itamPatchDestination, majorVersion,minorVersion, rollupVersion)

if(xflowProductUpdate=="xFlow"):
    if os.path.exists(xflowPatchDestination + "//xFlow//xFlowAnalyst//Binaries//"):
        xFlow = os.listdir(xflowPatchDestination + "//xFlow//xFlowAnalyst//Binaries//")
        for xFlowAnalystPatch in xFlow:
            logging.info("xFlow Binary is " + xFlowAnalystPatch)
    else:
        xFlowAnalystPatch = ""

    if os.path.exists(xflowPatchDestination + "//xFlow//CollaborationServer//Binaries//"):
        collabSrvr = os.listdir(xflowPatchDestination + "//xFlow//CollaborationServer//Binaries//")
        for collaborationPatch in collabSrvr:
            logging.info("CollabServer binary is " + collaborationPatch)
    else:
        collaborationPatch=""

    if os.path.exists(xflowPatchDestination + "//xFlow//SearchServer//Binaries//"):
        SrchSrvr = os.listdir(xflowPatchDestination + "//xFlow//SearchServer//Binaries//")
        for srchSrvrPatch in SrchSrvr:
            logging.info("SearchServer binary is " + srchSrvrPatch)
    else:
        srchSrvrPatch=""

logging.info("Deleting existing products folder...")
shutil.rmtree(commonInstallerDestination + "//products")

# Check for if installation is fresh
if(freshInstall=="Auto" or freshInstall=="ISO"):
    logging.info("Attempting to copy products folder from " + productInstallSource + " to " + commonInstallerDestination + "//products")
    shutil.copytree(productInstallSource,commonInstallerDestination + "//products")

if(freshInstall=="ISO"):
    logging.info("Creating ISO file now")
    os.chdir(commonInstallerDestination)
    command = "c:\\mkisofs -l -r -v -V CASM_" + majorVersion + "_" + minorVersion + "_WIN -o CASM_" + majorVersion + "_" + minorVersion + "_WIN.iso " + commonInstallerDestination
    logging.info("ISO Creation command : " + command)
    p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
    stdout, stderr = p.communicate()
    logging.info("ISO creation initiated with error code " + str(p.returncode))
    sys.exit()

# Updating response files

patchProps = load_properties(responseFilePath + "//PatchConfig.properties","=","#")
if(sdmProductUpdate =="SDM"):
    sdmLocalePatchNameList = patchProps['sdm.locale.patch.name.list']
    sdmSelectedPatchName= patchProps['sdm.selected.patch.name']
    sdmBinaryPatchNameList = patchProps['sdm.binary.patch.name.list']
    sdmLocalePatchName = patchProps['sdm.locale.patch.name']
    sdmBinaryPatchName = patchProps['sdm.binary.patch.name']

    updateResponseFiles(responseFilePath + "//PatchConfig.properties","sdm.locale.patch.name.list=" + sdmLocalePatchNameList,"sdm.locale.patch.name.list=" + sdmLocaleUpdate[:-4])
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","sdm.selected.patch.name=" + sdmSelectedPatchName,"sdm.selected.patch.name=" + patchVersion)
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","sdm.binary.patch.name.list=" + sdmBinaryPatchNameList,"sdm.binary.patch.name.list=" + sdmCumulativeUpdate[:-4])
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","sdm.locale.patch.name=" + sdmLocalePatchName,"sdm.locale.patch.name=" + sdmLocaleUpdate[:-4])
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","sdm.binary.patch.name=" + sdmBinaryPatchName,"sdm.binary.patch.name=" + sdmCumulativeUpdate[:-4])

if(xflowProductUpdate=="xFlow"):
    if(collaborationPatch != ""):
        collabServerSelectedPatchName = patchProps['collab.server.selected.patch.name']
        collabBinaryPatchNameList = patchProps['collab.binary.patch.name.list']
        collabBinaryPatchName = patchProps['collab.binary.patch.name']

        updateResponseFiles(responseFilePath + "//PatchConfig.properties","collab.server.selected.patch.name=" + collabServerSelectedPatchName, "collab.server.selected.patch.name=" + patchVersion)
        updateResponseFiles(responseFilePath + "//PatchConfig.properties","collab.binary.patch.name.list=" + collabBinaryPatchNameList,"collab.binary.patch.name.list=" + collaborationPatch[:-4])
        updateResponseFiles(responseFilePath + "//PatchConfig.properties","collab.binary.patch.name=" + collabBinaryPatchName,"collab.binary.patch.name=" + collaborationPatch[:-4])

if(xflowProductUpdate=="xFlow"):
    if(srchSrvrPatch != ""):
        searchServerSelectedPatchName = patchProps['search.server.selected.patch.name']
        searchServerPatchNameList = patchProps['searchserver.patch.name.list']
        searchServerPatchName = patchProps['searchserver.patch.name']

        updateResponseFiles(responseFilePath + "//PatchConfig.properties","search.server.selected.patch.name=" + searchServerSelectedPatchName, "search.server.selected.patch.name=" + patchVersion)
        updateResponseFiles(responseFilePath + "//PatchConfig.properties","searchserver.patch.name.list=" + searchServerPatchNameList,"searchserver.patch.name.list=" + srchSrvrPatch[:-4])
        updateResponseFiles(responseFilePath + "//PatchConfig.properties","searchserver.patch.name=" + searchServerPatchName,"searchserver.patch.name=" + srchSrvrPatch[:-4])

if(xflowProductUpdate=="xFlow"):
    if (xFlowAnalystPatch != ""):
        xflowPatchNameList = patchProps['xflow.patch.name.list']
        xflowPatchName = patchProps['xflow.patch.name']
        xflowSelectedPatchName = patchProps['xflow.selected.patch.name']

        updateResponseFiles(responseFilePath + "//PatchConfig.properties","xflow.patch.name.list=" + xflowPatchNameList,"xflow.patch.name.list=" + xFlowAnalystPatch[:-4])
        updateResponseFiles(responseFilePath + "//PatchConfig.properties", "xflow.patch.name=" + xflowPatchName,"xflow.patch.name=" + xFlowAnalystPatch[:-4])
        updateResponseFiles(responseFilePath + "//PatchConfig.properties","xflow.selected.patch.name=" + xflowSelectedPatchName,"xflow.selected.patch.name=" + patchVersion)

if(slcmProductUpdate=="SLCM"):
    slcmSelectedPatchName = patchProps['slcm.selected.patch.name']
    slcmBinaryPatchNameList = patchProps['slcm.binary.patch.name.list']
    slcmBinaryPatchName = patchProps['slcm.binary.patch.name']
    if(freshInstall!="Auto"):
        slcmPatchNumbers = patchProps['slcm.patch.numbers']

    updateResponseFiles(responseFilePath + "//PatchConfig.properties","slcm.selected.patch.name=" + slcmSelectedPatchName,"slcm.selected.patch.name=" + patchVersion)
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","slcm.binary.patch.name.list=" + slcmBinaryPatchNameList,"slcm.binary.patch.name.list=" + slcmCumulativeUpdate[:-4])
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","slcm.binary.patch.name=" + slcmBinaryPatchName,"slcm.binary.patch.name=" + slcmCumulativeUpdate[:-4])
    if (freshInstall != "Auto"):
        updateResponseFiles(responseFilePath + "//PatchConfig.properties","slcm.patch.numbers=" + slcmPatchNumbers,"slcm.patch.numbers=" + slcmCumulativeUpdate[:-4])

if(itamProductUpdate=="ITAM"):
    itamSelectedPatchName = patchProps['itam.selected.patch.name']
    itamPatchNameList = patchProps['itam.patch.name.list']
    itamPatchName = patchProps['itam.patch.name']

    updateResponseFiles(responseFilePath + "//PatchConfig.properties","itam.selected.patch.name=" + itamSelectedPatchName, "itam.selected.patch.name=" + patchVersion)
    updateResponseFiles(responseFilePath + "//PatchConfig.properties","itam.patch.name.list=" + itamPatchNameList,"itam.patch.name.list=" + itamCumulativeUpdate[:-4])
    updateResponseFiles(responseFilePath + "//PatchConfig.properties", "itam.patch.name=" + itamPatchName,"itam.patch.name=" + itamCumulativeUpdate[:-4])

if(ussProductUpdate=="SelfService"):
    ussSelectedPatchName = patchProps['uss.selected.patch.name']
    ussBinaryPatchNameList = patchProps['uss.binary.patch.name.list']
    ussBinaryPatchName = patchProps['uss.binary.patch.name']

    updateResponseFiles(responseFilePath + "//PatchConfig.properties","uss.selected.patch.name=" + ussSelectedPatchName, "uss.selected.patch.name=" + patchVersion)
    updateResponseFiles(responseFilePath + "//PatchConfig.properties", "uss.binary.patch.name.list=" + ussBinaryPatchNameList,"uss.binary.patch.name.list=" + ussCumulativeUpdate[:-4])
    updateResponseFiles(responseFilePath + "//PatchConfig.properties", "uss.binary.patch.name=" + ussBinaryPatchName,"uss.binary.patch.name=" + ussCumulativeUpdate[:-4])

logging.info("Completed updating response files....")

"""
logging.info("Deleting existing products folder...")
shutil.rmtree(commonInstallerDestination + "//products")

# Check for if installation is fresh
if(freshInstall=="Auto"):
    logging.info("Attempting to copy products folder from " + productInstallSource + " to " + commonInstallerDestination + "//products")
    shutil.copytree(productInstallSource,commonInstallerDestination + "//products")
"""
# Write patch values to file for Jenkins job
if(sdmProductUpdate=="SDM"):
    f= open(workspaceDir + "//SDM_Patch.txt","w+")
    f.write(sdmCumulativeUpdate)
    f.close()

if(xflowProductUpdate=="xFlow"):
    f= open(workspaceDir + "//xFlow_Patch.txt","w+")
    f.write(xflowCumulativeUpdate)
    f.close()

if(ussProductUpdate=="SelfService"):
    f= open(workspaceDir + "//USS_Patch.txt","w+")
    f.write(ussCumulativeUpdate)
    f.close()

if(slcmProductUpdate=="SLCM"):
    f= open(workspaceDir + "//SLCM_Patch.txt","w+")
    f.write(slcmCumulativeUpdate)
    f.close()

if(itamProductUpdate=="ITAM"):
    f= open(workspaceDir + "//ITAM_Patch.txt","w+")
    f.write(itamCumulativeUpdate)
    f.close()

f= open(workspaceDir + "//CI_Patch.txt","w+")
f.write(commonInstallerFile)
f.close()

# Calling setup
logging.info("Calling Common Installer setup to initiate the upgrade now.....")
logging.info("Please review the install log file install.log under CASM Folder")

os.chdir(commonInstallerDestination)
command = "setup -S " + responseFilePath
p = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
stdout, stderr = p.communicate()
logging.info("Setup initiated with error code " + str(p.returncode))





