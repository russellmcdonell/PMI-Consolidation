
# pylint: disable=line-too-long

'''
A script to read in the secondary PMI file for patients without a alternate UR number.
Then read in the master PMI file and check each record looking for any secondary PMI records that match.

SYNOPSIS
$ python findUR.py masterDirectory [-e masterExtractDirectory|--masterExtractDir=masterExtractDirectory]
                   secondaryDirectory [-f secondaryExtractDirectory|--secondaryExtractDir=secondaryExtractDirectory] [-E|--Extensive]
                   [-m masterDebugKey|--masterDebugKey=masterDebugKey] [-n masterDebugCount|--masterDebugCount=masterDebugCount]
                   [-s secondaryDebugKey|--secondaryDebugKey=secondaryDebugKey] [-t secondaryDebugCount|--secondaryDebugCount=secondaryDebugCount]
                   [-v loggingLevel|--verbose=loggingLeve] [-o logfile|--logfile=logfile]


OPTIONS
masterDirectory
The directory containing the master configuration (master.cfg) plus the master PMI file (master.csv)

secondaryDirectory
The directory containing the secondary configuration (secondary.cfg) plus the secondary PMI file (secondary.csv)

-e masterExtractDir|--masterExtractDir=masterExtractDir
The optional extract sub-directory, of the master directory, containing the extract master CSV file and parameters.

-f secondaryExtractDir|--secondaryExtractDir=secondaryExtractDir
The optional extract sub-directory, of the secondary directory, containing the extract secondary CSV file and parameters.

-E|--Extensive
Invoke extensive checking for possible duplicates. Extensive checking can invoke various function on the core data elements (FamilyName, GivenName, Birthdate, Sex)
plus simple checks of equality for any other field in the master and secondary PMI files.

-m masterDebugKey|--masterDebugKey=masterDebugKey
The key for triggering logging of information about a specific master record. Default is None

-n masterDebugCount|--masterDebugCount=masterDebugCount
A counter to trigger progress logging; a progress message is created every masterDebugCount(th) master record. Default is 50000

-s secondaryDebugKey|--secondaryDebugKey=secondaryDebugKey
The key for triggering logging of information about a specific secondary record. Default is None

-t secondaryDebugCount|--secondaryDebugCount=secondaryDebugCount
A counter to trigger progress logging; a progress message is created every secondaryDebugCount(th) secondary record. Default is 50000

-v loggingLevel|--verbose=loggingLevel
Set the level of logging that you want.

-o logfile|--logfile=logfile
The name of a log file where you want all messages captured.


THE MAIN CODE
Start by parsing the command line arguements and setting up logging.

Then read in the Secondary PMI file and create a "key" for each patient that does not have a alternate UR number.
(If the secondary PMI file contains deleted, merged or alias patients, then we ignore them)
That "key" will be made up of the cleaned up family name, cleaned up given name, birthdate and sex of the patient.

Next read in the master PMI file and compute the same "key" for each patient.
The Master PMI file can contains aliases and merged patients.
We do look for both full and partial matches against aliases and we report the match against the alias as we assume the alias and the
primary patient are a perfect match. We already know that they have the same UR number, as this has been checked in checkMaster.pl
If the Extensive checking option is invoked, then compute an overall confidence score good matches based upon the extended matching fields.

Determine which patients in the Master PMI file are records of interest. Then re-read the Master PMI file and save the information about patient of interest.

Finally, re-read the secondary PMI file and printout the matching information about patients of interest from the master PMI file

'''

# pylint: disable=invalid-name, bare-except, pointless-string-statement, unspecified-encoding, too-many-lines

import os
import sys
import csv
import argparse
import logging
import re
from importlib.machinery import SourceFileLoader
from importlib.util import spec_from_loader, module_from_spec
from openpyxl.styles import NamedStyle
import data as d
import functions as f

# This next section is plagurised from /usr/include/sysexits.h
EX_OK = 0        # successful termination
EX_WARN = 1        # non-fatal termination with warnings

EX_USAGE = 64        # command line usage error
EX_DATAERR = 65        # data format error
EX_NOINPUT = 66        # cannot open input
EX_NOUSER = 67        # addressee unknown
EX_NOHOST = 68        # host name unknown
EX_UNAVAILABLE = 69    # service unavailable
EX_SOFTWARE = 70    # internal software error
EX_OSERR = 71        # system error (e.g., can't fork)
EX_OSFILE = 72        # critical OS file missing
EX_CANTCREAT = 73    # can't create (user) output file
EX_IOERR = 74        # input/output error
EX_TEMPFAIL = 75    # temp failure; user is invited to retry
EX_PROTOCOL = 76    # remote error in protocol
EX_NOPERM = 77        # permission denied
EX_CONFIG = 78        # configuration error


if __name__ == '__main__':
    '''
The main code
Start by parsing the command line arguements and setting up logging.
Then compute 'keys' for secondary PMI records that have an AltUR. Then compute the same 'keys' for master PMI records with a UR that matches an Alt UR in the secondary PMI records.
Check the quality of the match and remember the master PMI data. Then re-read the secondary PMI file and print matching information.
    '''

    # Save the program name
    d.progName = sys.argv[0]
    d.progName = d.progName[0:-3]        # Strip off the .py ending
    d.scriptType = 'find'

    # Get the options
    parser = argparse.ArgumentParser(description='Look for matches between the secondary PMI file and the master PMI file, where AltUR is missing or incorrect')
    parser.add_argument ('masterDir', metavar='masterDirectory', help='The name of directory containg the master configuration and master PMI file (master.csv)')
    parser.add_argument ('secondaryDir', metavar='secondaryDirectory', help='The name of directory containg the secondary configuration secondary PMI file (secondary.csv)')
    parser.add_argument ('-e', '--masterExtractDir', dest='masterExtractDir', metavar='masterExtractDirectory', help='The name of the master directory sub-directory that contains the extract master CSV file and configuration specific to the extract')
    parser.add_argument ('-f', '--secondaryExtractDir', dest='secondaryExtractDir', metavar='secondaryExtractDirectory', help='The name of the secondary directory sub-directory that contains the extract secondary CSV file and configuration specific to the extract')
    parser.add_argument ('-E', '--Extensive', dest='Extensive', action='store_true', help='Invoke Extensive checking to look for possible matches')
    parser.add_argument ('-m', '--masterDebugKey', dest='masterDebugKey', metavar='masterDebugKey', help='The key for triggering logging of information about a specific master record')
    parser.add_argument ('-n', '--masterDebugCount', dest='masterDebugCount', metavar='masterDebugCount', type=int, default=50000, help='A counter to trigger progress logging; a message every masterDebugCount(th) master record')
    parser.add_argument ('-s', '--secondaryDebugKey', dest='secondaryDebugKey', metavar='secondaryDebugKey', help='The key for triggering logging of information about a specific secondary record')
    parser.add_argument ('-t', '--secondaryDebugCount', dest='secondaryDebugCount', metavar='secondaryDebugCount', type=int, default=50000, help='A counter to trigger progress logging; a message every secondaryDebugCount(th) secondary record')
    parser.add_argument ('-v', '--verbose', dest='verbose', type=int, choices=range(0,5), help='The level of logging\n\t0=CRITICAL,1=ERROR,2=WARNING,3=INFO,4=DEBUG')
    parser.add_argument ('-o', '--logfile', dest='logfile', metavar='logfile', help='The name of a logging file')
    args = parser.parse_args()

    logging_levels = {0:logging.CRITICAL, 1:logging.ERROR, 2:logging.WARNING, 3:logging.INFO, 4:logging.DEBUG}
    logfmt = d.progName + ' [%(asctime)s]: %(message)s'
    if args.verbose:    # Change the logging level from "WARN" if the -v vebose option is specified
        loggingLevel = args.verbose
        if args.logfile :        # and send it to a file if the -o logfile option is specified
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', level=logging_levels[loggingLevel], filename=args.logfile)
        else:
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', level=logging_levels[loggingLevel])
    else:
        if args.logfile :        # send the default (WARN) logging to a file if the -o logfile option is specified
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', filename=args.logfile)
        else:
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p')


    d.masterDir = args.masterDir
    d.secondaryDir = args.secondaryDir
    d.masterExtractDir = args.masterExtractDir
    d.secondaryExtractDir = args.secondaryExtractDir
    d.Extensive = args.Extensive
    d.masterDebugKey = args.masterDebugKey
    d.masterDebugCount = args.masterDebugCount
    d.secondaryDebugKey = args.secondaryDebugKey
    d.secondaryDebugCount = args.secondaryDebugCount


    # Read in the extract configuration file if required
    d.masterReportingColumns = []
    f.getMasterConfig(False)
    if d.masterExtractDir:
        # Read in the extract configuration file
        f.getMasterConfig(True)

    # Import in the masterDir/cleanMaster%.py and masterDir/linkMaster.py - the master PMI extract specific code
    if not os.path.exists(f'./{d.masterDir}/cleanMaster.py'):
        logging.fatal('./%s/cleanMaster.py not found', d.masterDir)
        sys.exit(1)
    try:
        loader = SourceFileLoader('mc', f'./{d.masterDir}/cleanMaster.py')
        spec = spec_from_loader('mc', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.mc = module
    except:
        logging.fatal('importing ./%s/cleanMaster.py failed', d.masterDir)
        sys.exit(1)
    if not os.path.exists(f'./{d.masterDir}/linkMaster.py'):
        logging.fatal('./%s/linkMaster.py not found', d.masterDir)
        sys.exit(1)
    try:
        loader = SourceFileLoader('ml', f'./{d.masterDir}/linkMaster.py')
        spec = spec_from_loader('ml', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.ml = module
    except:
        logging.fatal('importing ./%s/linkMaster.py failed', d.masterDir)
        sys.exit(1)

    # Read in the secondary configuration file
    d.secondaryReportingColumns = []
    f.getSecondaryConfig(False)

    # Read in the extract configuration file if required
    if d.secondaryExtractDir:
        # Read in the extract configuration file
        f.getSecondaryConfig(True)

    # Assemble the reporting columns
    d.reportingColumns = []
    d.reportingDates = []
    wantDate = re.compile('date', flags=re.IGNORECASE)
    # Start with things in both
    for col in d.masterReportingColumns:
        if col in d.secondaryReportingColumns:
            d.reportingColumns.append(col)
            if wantDate.search(col) is not None:
                d.reportingDates.append(col)
    # Add things in master only
    for col in d.masterReportingColumns:
        if col not in d.secondaryReportingColumns:
            d.reportingColumns.append(col)
            if wantDate.search(col) is not None:
                d.reportingDates.append(col)
    # Add things in secondary only
    for col in d.secondaryReportingColumns:
        if col not in d.masterReportingColumns:
            d.reportingColumns.append(col)
            if wantDate.search(col) is not None:
                d.reportingDates.append(col)


    # Import in the secondaryDir/cleanSecondary%.py and secondaryDir/linkSecondary.py - the secondary PMI extract specific code
    if not os.path.exists(f'./{d.secondaryDir}/cleanSecondary.py'):
        logging.fatal('./%s/cleanSecondary.py not found', d.secondaryDir)
        sys.exit(1)
    try:
        loader = SourceFileLoader('sc', f'./{d.secondaryDir}/cleanSecondary.py')
        spec = spec_from_loader('sc', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.sc = module
    except:
        logging.fatal('importing ./%s/cleanSecondary.py failed', d.secondaryDir)
        sys.exit(1)
    if not os.path.exists(f'./{d.secondaryDir}/linkSecondary.py'):
        logging.fatal('./%s/linkSecondary.py not found', d.secondaryDir)
        sys.exit(1)
    try:
        loader = SourceFileLoader('sl', f'./{d.secondaryDir}/linkSecondary.py')
        spec = spec_from_loader('sl', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.sl = module
    except:
        logging.fatal('importing ./%s/linkSecondary.py failed', d.secondaryDir)
        sys.exit(1)

    # Open Error file
    f.openErrorFile()

    # Read in the Found PIDs and altURs
    f.getFound()

    # Read in the Not Found PIDs - some records might might have a valid Alt URs
    f.getNotFound()

    # Read in the Matched PIDs - matches found and no checking required
    f.getMatched()

    # Read in the Not Matched PIDs - confirmed non-matches - further checking requried
    f.getNotMatched()


    # Open the notFoundDone file
    d.date_style = NamedStyle(name='Date', number_format='dd-mmm-yyyy')
    f.PrintHeading('nf', 2, 'Checked', 'Message')

    # Pass 1 - pick out the keys from the secondary PMI file for records that could be matched [those without an AltUR]
    # Open the cleaned up CSV secondary PMI file
    notFounddn = 0        # Count of secondary records that have not been found and are known to be not in the master PMI [notFound.xlsx]
    d.fullKey = {}            # The Full Key / secondary PMI the record number(s) for this key
    d.keySsx = {}            # The Sounds Like (Soundex) Keys and the record number for this key
    d.keySdm = {}            # The Sounds Like (Double Metaphone) Keys and the record number for this key
    d.keySny = {}            # The Sounds Like (NYSIIS) Keys and the record number for this key
    d.key123 = {}            # The Family Name, Sex and DOB Key and the rec. number
    d.key124 = {}            # The Family Name, Sex and Given Name Key and the rec. no.
    d.key134 = {}            # The Family Name, DOB and Given Name Key and the rec.rd no.
    d.key234 = {}            # The Sex, DOB and Given Name Key and the record number
    secondaryCSV = None
    if d.secondaryExtractDir:
        secondaryCSV = f'./{d.secondaryDir}/{d.secondaryExtractDir}/secondary.csv'
    else:
        secondaryCSV = f'./{d.secondaryDir}/secondary.csv'
    with open(secondaryCSV, 'rt') as csvfile:
        secondaryPMI = csv.reader(csvfile, dialect='excel')
        d.secondaryRecNo = 0
        heading = True
        for d.csvfields in secondaryPMI:
            if heading:
                heading = False
                continue

            # Report progress
            d.secondaryRecNo += 1
            if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                logging.info('%d secondary PMI records read', d.secondaryRecNo)

            # Only check patients without an altUR number and those where the altUR is know to be wrong (notMatch.xlsx)
            altUR = d.sc.secondaryCleanAltUR()    # Get the altUR number
            pid = d.sc.secondaryCleanPID()        # Get the PID number
            if pid in d.notMatched :        # Include this record if the AltUR in this record is known to be wrong
                altUR = ''
            if altUR != '':
                continue

            # Don't check patients where the patient is know to be not in the master PMI (notFound.xlsx)
            if pid in d.notFound :        # Report this record as not found
                f.PrintSecondary(True, 'Y', 2, 'nf', f'not found in {d.masterLongName}', 0)
                notFounddn += 1
                continue

            # Check if this pid has already been found (i.e. the pid and altUR are in found.xlsx because the record has been previously found)
            if pid in d.foundUR:
                d.foundSecondaryRec[pid] = d.secondaryRecNo    # Secondary PMI record number for a secondary PMI PID
                d.secondaryFoundPID[d.secondaryRecNo] = pid
                d.recStatus[d.secondaryRecNo] = 6
                d.foundRec[d.secondaryRecNo] = ''        # The matching master PMI record for this secondary PMI record (-1 == none found yet/to be assigned)
                d.extras[d.secondaryRecNo] = ''
                continue


            # Clean up the Secondary PMI family name, given name name and sex
            Sf = d.sc.secondaryCleanFamilyName()
            Sg = d.sc.secondaryCleanGivenName()
            Sdob = d.sc.secondaryCleanDOB()
            Ssex = d.sc.secondaryCleanSex()
            thisKey = Sf + '~' + Sg + '~' + Sdob + '~' + Ssex
            if (d.secondaryDebugKey) and (d.secondaryDebugKey in thisKey):
                logging.info('%s Test Patient:%s:%s:%s', d.secondaryLongName, pid, altUR, thisKey)

            # This is the full key.
            # Save it, and the fact that the status of this secondary records is 'unknown'
            if thisKey not in d.fullKey:
                d.fullKey[thisKey] = []
            d.fullKey[thisKey].append(d.secondaryRecNo)
            d.recStatus[d.secondaryRecNo] = -1
            d.foundRec[d.secondaryRecNo] = ''
            d.extras[d.secondaryRecNo] = ''

            # Now compute the sounds key
            soundKey = f.Sounds(Sf, Sg)
            (Sfny, Sfdm, Sfsx, Sgny, Sgdm, Sgsx) = soundKey.split('~')
            soundKeysx = Sfsx + '~' + Sgsx + '~' + Sdob + '~' + Ssex
            if soundKeysx not in d.keySsx:
                d.keySsx[soundKeysx] = []
            d.keySsx[soundKeysx].append(d.secondaryRecNo)
            soundKeydm = Sfdm + '~' + Sgdm + '~' + Sdob + '~' + Ssex
            if soundKeydm not in d.keySdm:
                d.keySdm[soundKeydm] = []
            d.keySdm[soundKeydm].append(d.secondaryRecNo)
            soundKeyny = Sfny + '~' + Sgny + '~' + Sdob + '~' + Ssex
            if soundKeyny not in d.keySny:
                d.keySny[soundKeyny] = []
            d.keySny[soundKeyny].append(d.secondaryRecNo)

            # And finally the four partial keys
            thisKey = Sf + '~' + Sg + '~' + Sdob
            if thisKey not in d.key123:
                d.key123[thisKey] = []
            d.key123[thisKey].append(d.secondaryRecNo)
            thisKey = Sf + '~' + Sg + '~' + Ssex
            if thisKey not in d.key124:
                d.key124[thisKey] = []
            d.key124[thisKey].append(d.secondaryRecNo)
            thisKey = Sf + '~' + Sdob + '~' + Ssex
            if thisKey not in d.key134:
                d.key134[thisKey] = []
            d.key134[thisKey].append(d.secondaryRecNo)
            thisKey = Sg + '~' + Sdob + '~' + Ssex
            if thisKey not in d.key234:
                d.key234[thisKey] = []
            d.key234[thisKey].append(d.secondaryRecNo)

            if d.Extensive and (pid not in d.foundSecondaryRec) :        # Collect and pack Extensive checking data (if required)
                d.ExtensiveSecondaryRecKey[d.secondaryRecNo] = f.Sounds(Sf, Sg)
                if Sdob == '':
                    d.ExtensiveSecondaryBirthdate[d.secondaryRecNo] = d.futureBirthdate
                else:
                    d.ExtensiveSecondaryBirthdate[d.secondaryRecNo] = f.Birthdate(Sdob)
                if d.useMiddleNames:
                    d.ExtensiveSecondaryMiddleNames[d.secondaryRecNo] = f.secondaryField('MiddleNames').upper()
                thisHash = {}
                for field in (sorted(d.ExtensiveFields.keys())):
                    fieldData = f.secondaryField(field)
                    if fieldData == '':
                        thisHash[field] = None
                    else:
                        thisHash[field] = hash(fieldData)
                d.ExtensiveOtherSecondaryFields[d.secondaryRecNo] = thisHash

    if d.secondaryExtractDir:
        f.PrintClose('nf', 0, 5, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_NotFound_Done.xlsx')
    else:
        f.PrintClose('nf', 0, 5, f'{d.secondaryDir}/{d.secondaryShortName}_NotFound_Done.xlsx')

    logging.info('End of Pass 1')



    # Pass 2 - read the Master PMI file and match keys updating status of records in the Secondary PMI file as you go
    #       Also, save the merged patients links as you will need them later
    #       And build up the data about probabalistic matched (Possible Finds)
    # Open the cleaned up CSV master PMI file
    d.possExtensiveFinds = {}                # all the matches for each of the matched secondary PMI record
    masterCSV = None
    if d.masterExtractDir:
        masterCSV = f'./{d.masterDir}/{d.masterExtractDir}/master.csv'
    else:
        masterCSV = f'./{d.masterDir}/master.csv'
    with open(masterCSV, 'rt') as csvfile:
        masterPMI = csv.reader(csvfile, dialect='excel')
        d.masterRecNo = 0
        heading = True
        for d.csvfields in masterPMI :            # Check every Master PMI file record
            if heading:
                heading = False
                continue

            # Report progress
            d.masterRecNo += 1
            if (d.masterRecNo % d.masterDebugCount) == 0:
                logging.info('%d master PMI records read', d.masterRecNo)

            # Save alias and merge links for this record
            f.masterSaveLinks()

            # Save the alias and merge patient information
            if d.ml.masterIsAlias():
                f.masterSetAlias()
            if d.ml.masterIsMerged():
                f.masterSetMerged()

            # Check for Found patients
            ur = d.mc.masterCleanUR()        # The UR number

            # Check for secondary PIDs that have already been found (this UR is in found.xlsx)
            if (not d.ml.masterIsAlias()) and (not d.ml.masterIsMerged()) and (ur in d.foundPID):
                secondaryPIDs = d.foundPID[ur].split('~')            # The list of secondary PIDs with AltURs that match this master UR
                for secondaryPID in (secondaryPIDs) :                # Check that each of these secondary PIDs was found in the secondary PMI (see above)
                    if secondaryPID in d.foundSecondaryRec:
                        secRecNo = d.foundSecondaryRec[secondaryPID]
                        f.SaveStatus(secRecNo, d.recStatus[secRecNo], '')
                    else:
                        d.feCSV.writerow([f'{d.progName}:ERROR in found.xlsx:{ur},{d.foundPID[ur]} - {d.secondaryLongName} {d.secondaryPIDname} {secondaryPID} not found'])


            # Clean up the Master PMI family name and given name name
            Mf = d.mc.masterCleanFamilyName()
            Mg = d.mc.masterCleanGivenName()
            Mdob = d.mc.masterCleanDOB()
            Msex = d.mc.masterCleanSex()
            testKey = False
            thisKey = Mf + '~' + Mg + '~' + Mdob + '~' + Msex
            if (d.masterDebugKey) and (d.masterDebugKey == thisKey):
                logging.info('%s Test Patient:%s:%s:%s', d.masterLongName, pid, ur, thisKey)
                testKey = True


            # Check if this Master PMI file record matches any Secondary PMI records
            Mfny = Mfdm1 = Mfdm2 = Mfsx = Mgny = Mgdm = Mgsx = My = Mm = Md = masterBirthdate = Mmn = None
            soundKey = f.Sounds(Mf, Mg)
            (Mfny, Mfdm, Mfsx, Mgny, Mgdm, Mgsx) = soundKey.split('~')
            found = False
            if thisKey in d.fullKey:
                found = True
                for secRecNo in d.fullKey[thisKey]:
                    f.SaveStatus(secRecNo, 6, '')

            # Check for a Sound match
            if not found:
                soundFound = ''
                soundKeydm = Mfdm + '~' + Mgdm+ '~' + Mdob + '~' + Msex
                if soundKeydm in d.keySdm:
                    soundFound = '1'
                else:
                    soundFound = '0'
                soundKeyny = Mfny + '~' + Mgny + '~' + Mdob + '~' + Msex
                if soundKeyny in d.keySny:
                    soundFound += '1'
                else:
                    soundFound += '0'
                soundKeysx = Mfsx + '~' + Mgsx + '~' + Mdob + '~' + Msex
                if soundKeysx in d.keySsx:
                    soundFound += '1'
                else:
                    soundFound += '0'
                if soundFound != '000':
                    found = True
                    if soundKeydm in d.keySdm:
                        for secRecNo in d.keySdm[soundKeydm]:
                            f.SaveStatus(secRecNo, 5, soundFound)
                    elif soundKeyny in d.keySny:
                        for secRecNo in d.keySny[soundKeyny]:
                            f.SaveStatus(secRecNo, 5, soundFound)
                    elif soundKeysx in d.keySsx:
                        for secRecNo in d.keySsx[soundKeysx]:
                            f.SaveStatus(secRecNo, 5, soundFound)

            # And finally the four partial keys
            if not found:
                thisKey = Mf + '~' + Mg + '~' + Mdob
                if thisKey in d.key123:
                    found = True
                    for secRecNo in d.key123[thisKey]:
                        f.SaveStatus(secRecNo, 4, '')
            if not found:
                thisKey = Mf + '~' + Mg + '~' + Msex
                if thisKey in d.key124:
                    found = True
                    for secRecNo in d.key124[thisKey]:
                        f.SaveStatus(secRecNo, 3, '')
            if not found:
                thisKey = Mf + '~' + Mdob + '~' + Msex
                if thisKey in d.key134:
                    found = True
                    for secRecNo in d.key134[thisKey]:
                        f.SaveStatus(secRecNo, 2, '')
            if not found:
                thisKey = Mg + '~' + Mdob + '~' + Msex
                if thisKey in d.key234:
                    found = True
                    for secRecNo in d.key234[thisKey]:
                        f.SaveStatus(secRecNo, 1, '')

            # For Extensive checking we compute a confidence level that this master record matches each secondary record
            # There can be multiple secondary records claiming to be linked to each master record
            if d.Extensive:
                if Mdob == '':
                    masterBirthdate = d.futureBirthdate
                else:
                    masterBirthdate = f.Birthdate(Mdob)
                if d.useMiddleNames:
                    Mmn = f.masterField('MiddleNames').upper()
                masterOtherFields = {}
                for field in (sorted(d.ExtensiveFields.keys())):
                    fieldData = f.masterField(field)
                    if fieldData == '':
                        masterOtherFields[field] = None
                    else:
                        masterOtherFields[field] = hash(fieldData)
                for secRecNo, stringKey in d.ExtensiveSecondaryRecKey.items():
                    # Unpack the secondary PMI items for this secondary PMI record
                    (Sfny, Sfdm, Sfsx, Sgny, Sgdm, Sgsx) = stringKey.split('~')
                    secondaryBirthdate = d.ExtensiveSecondaryBirthdate[secRecNo]
                    Smn = None
                    if d.useMiddleNames:
                        Smn = d.ExtensiveSecondaryMiddleNames[secRecNo]
                    secondaryOtherFields = d.ExtensiveOtherSecondaryFields[secRecNo]
                    weight = 1.0
                    if Mf == Sf:
                        soundFamilyNameConfidence = 100.0
                    else:
                        (soundFamilyNameConfidence, weight) = f.FamilyNameSoundCheck(Mf, Mfny, Mfdm, Mfsx, Sf, Sfny, Sfdm, Sfsx, 1.0)
                    if Mg == Sg:
                        soundGivenNameConfidence = 100.0
                    else:
                        (soundGivenNameConfidence, weight) = f.GivenNameSoundCheck(Mg, Mgny, Mgdm, Mgsx, Sg, Sgny, Sgdm, Sgsx, 1.0)

                    # Compute the goodness of fit between this secondary record and the master record
                    totalConfidence = 0
                    totalWeight = 0
                    # Start with the things that require algorithms
                    for coreRoutine, (thisWeight, thisParam) in d.ExtensiveRoutines.items():
                        weight = 0.0
                        confidence = 0.0
                        if coreRoutine == 'FamilyName':
                            (confidence, weight) = f.FamilyNameCheck(Mf, Sf, thisWeight)
                        elif coreRoutine == 'FamilyNameSound':
                            (confidence, weight) = f.FamilyNameSoundCheck(Mf, Mfny, Mfdm, Mfsx, Sf, Sfny, Sfdm, Sfsx, thisWeight)
                        elif coreRoutine == 'GivenName':
                            (confidence, weight) = f.GivenNameCheck(Mg, Sg, thisWeight)
                        elif coreRoutine == 'GivenNameSound':
                            (confidence, weight) = f.GivenNameSoundCheck(Mg, Mgny, Mgdm, Mgsx, Sg, Sgny, Sgdm, Sgsx, thisWeight)
                        elif coreRoutine == 'MiddleNames':
                            if d.useMiddleNames:
                                (confidence, weight) = f.MiddleNamesCheck(Mmn, Smn, thisWeight)
                        elif coreRoutine == 'MiddleNamesInitial':
                            if d.useMiddleNames:
                                (confidence, weight) = f.MiddleNamesInitialCheck(Mmn, Smn, thisWeight)
                        elif coreRoutine == 'Sex':
                            (confidence, weight) = f.SexCheck(Msex, Ssex, thisWeight)
                        elif coreRoutine == 'Birthdate':
                            (confidence, weight) = f.BirthdateCheck(Mdob, Sdob, thisWeight)
                        elif coreRoutine == 'BirthdateNearYear':
                            (confidence, weight) = f.BirthdateNearYearCheck(Mdob, Sdob, thisParam, thisWeight)
                        elif coreRoutine == 'BirthdateNearMonth':
                            (confidence, weight) = f.BirthdateNearMonthCheck(Mdob, Sdob, thisParam, thisWeight)
                        elif coreRoutine == 'BirthdateNearDay':
                            (confidence, weight) = f.BirthdateNearDayCheck(masterBirthdate, secondaryBirthdate, thisParam, thisWeight)
                        elif coreRoutine == 'BirthdateYearSwap':
                            (confidence, weight) = f.BirthdateYearSwapCheck(Mdob, Sdob, thisWeight)
                        elif coreRoutine == 'BirthdateDayMonthSwap':
                            (confidence, weight) = f.BirthdateDayMonthSwapCheck(Mdob, Sdob, thisWeight)
                        if weight > 0:
                            totalConfidence += confidence * weight
                            totalWeight += weight
                    # Then the things that just need matching (best done on hash values)
                    for field in (sorted(d.ExtensiveFields.keys())):
                        weight = d.ExtensiveFields[field]
                        if weight > 0:
                            if masterOtherFields[field] and secondaryOtherFields[field]:
                                if masterOtherFields[field] == secondaryOtherFields[field]:
                                    totalConfidence += 100.0 * weight
                                    totalWeight += weight
                                else:
                                    totalConfidence += 0.0 * weight
                                    totalWeight += weight
                    # Compute the total weight
                    if totalWeight > 0:
                        totalConfidence = totalConfidence / totalWeight
                    # And save this master data against this secondary record if we have an adequate match
                    # A master record can be matched to multiple secondary records, with varying degrees of confidence (multiple secondary rows with the same AltUR)
                    # A secondary record can only be matched to multiple master records if there are multiple master records with the same UR!!!
                    if totalConfidence >= d.ExtensiveConfidence:
                        if secRecNo not in d.possExtensiveFinds:
                            d.possExtensiveFinds[secRecNo] = {}
                        if totalConfidence not in d.possExtensiveFinds[secRecNo]:
                            d.possExtensiveFinds[secRecNo][totalConfidence] = []
                        d.possExtensiveFinds[secRecNo][totalConfidence].append([d.masterRecNo, soundFamilyNameConfidence, soundGivenNameConfidence])

    logging.info('End of Pass 2')


    # Find the links for aliases and merged patients
    f.masterFindAliases()
    f.masterFindMerged()


    # Pass 3 - Identify the Master PMI file records of interest
    d.secondaryRecNo = 0
    d.wantedMasterRec = {}
    for secRecNo, thisStatus in d.recStatus.items():

        # Report progress
        d.secondaryRecNo += 1
        if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
            logging.info('%d master PMI records records of interest identified', d.secondaryRecNo)

        if thisStatus >= 0:
            masterRecs = d.foundRec[secRecNo].split('~')
            for i, masterRecNo in enumerate(masterRecs):
                masterRecNo = int(masterRecNo)
                if masterRecNo == -1:
                    d.feCSV.writerow([f'{d.progName}:ERROR in found.xlsx:{d.secondaryFoundPID[secRecNo]},{d.foundUR[d.secondaryFoundPID[secRecNo]]} - {d.masterLongName} {d.masterURname} number not found'])
                else:
                    d.wantedMasterRec[masterRecNo] = True
                    # Check if an alias or merged patient and if so get master record as well
                    if masterRecNo in d.masterPrimRec:
                        d.wantedMasterRec[d.masterPrimRec[masterRecNo]] = True
                    if masterRecNo in d.masterNewRec:
                        d.wantedMasterRec[d.masterNewRec[masterRecNo]] = True

    for secRecNo, confidences in d.possExtensiveFinds.items():
        for confidence in confidences:
            for i, masterRecNos in enumerate(confidence):
                masterRecNo = masterRecNos[0]
                if masterRecNo not in d.wantedMasterRec:
                    # Report progress
                    d.secondaryRecNo += 1
                    if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                        logging.info('%d master PMI records records of interest identified', d.secondaryRecNo)
                    # Mark as wanted
                    d.wantedMasterRec[masterRecNo] = True
                    # Check if an alias or merged patient and if so get master record as well
                    if masterRecNo in d.masterPrimRec:
                        d.wantedMasterRec[d.masterPrimRec[masterRecNo]] = True
                    if masterRecNo in d.masterNewRec:
                        d.wantedMasterRec[d.masterNewRec[masterRecNo]] = True
    logging.info('End of Pass 3')



    # Pass 4 - re-read the Master PMI file, saving data of patients of interest
    with open(masterCSV, 'rt') as csvfile:
        masterPMI = csv.reader(csvfile, dialect='excel')
        d.masterRecNo = 0
        heading = True
        for d.csvfields in masterPMI:
            if heading:
                heading = False
                continue

            # Report progress
            d.masterRecNo += 1
            if (d.masterRecNo % d.masterDebugCount) == 0:
                logging.info('%d master PMI records re-read', d.masterRecNo)

            if d.masterRecNo in d.wantedMasterRec:
                f.masterSaveDetails()

    logging.info('End of Pass 4')



    # Pass 5 - re-read the secondary PMI file and printout the findings

    d.foundDoneVolume = 1
    d.foundToDoVolume = 1
    f.PrintHeading('nf', 2, 'Checked', 'Message')
    f.PrintHeading('fm', 3, 'Checked', 'Message')
    f.PrintHeading('fa', 3, 'Checked', 'Message')
    for thisFile in ['f', 'df', 'dpf', 'pfsnd', 'pfnf', 'pfng', 'pfnbd', 'pfnsx']:
        f.PrintHeading(thisFile + 'td', 3, 'Checked', 'Message')
        f.PrintHeading(thisFile + 'dn', 3, 'Checked', 'Message')
    f.PrintHeading('ef', 4, 'Checked', 'Identical based upon Extensive checking')
    f.PrintHeading('pf', 4, 'Checked', 'Similar based upon Extensive checking')

    d.mfound = 0        # Count of secondary records found to be a merged master PMI patent
    d.afound = 0        # Count of secondary records found to be an alias master PMI patent
    count = 0        # Count of secondary records in the secondary PMI file
    withoutAltUR = 0    # Count secondary records without altUR
    notFoundtd = 0        # Count of secondary records that have not been found in the master PMI and need to be checked to see if they are definitely not in the master PMI
    d.foundtd = 0        # Count of secondary records found
    d.founddn = 0        # Count of secondary records found (Done)
    d.dfoundtd = 0        # Count of secondary records with duplicate master records found
    d.dfounddn = 0        # Count of secondary records with duplicate master records found (Done)
    d.dpfoundtd = 0        # Count of secondary records with duplicate similar master records found
    d.dpfounddn = 0        # Count of secondary records with duplicate similar master records found (Done)
    d.extensivefind = 0    # Count of identical secondary records found using extensive matching
    d.probablefind = 0    # Count of similar secondary records found using extensive matching
    nmatch = 0
    notFound = 0
    dmatch = 0
    dmatchdn = 0
    pfoundsnd = 0
    pfoundsnddn = 0
    pfoundf = 0
    pfoundfdn = 0
    pfoundg = 0
    pfoundgdn = 0
    pfoundsx = 0
    pfoundsxdn = 0
    pfoundbd = 0
    pfoundbddn = 0
    with open(secondaryCSV, 'rt') as csvfile:
        secondaryPMI = csv.reader(csvfile, dialect='excel')
        d.secondaryRecNo = 0
        heading = True
        for d.csvfields in secondaryPMI:
            if heading:
                heading = False
                continue

            # Report progress
            d.secondaryRecNo += 1
            if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                logging.info('%d secondary PMI records re-read', d.secondaryRecNo)

            count += 1

            # Only report patients without an altUR number and those where the altUR is know to be wrong (notMatch.xlsx)
            altUR = d.sc.secondaryCleanAltUR()    # Get the altUR number
            pid = d.sc.secondaryCleanPID()        # Get the PID number
            if pid in d.notMatched :        # Include this record if the AltUR in this record is known to be wrong
                altUR = ''
            if altUR != '':
                continue

            # Don't check patients where the patient is know to be not in the master PMI (notFound.xlsx)
            if pid in d.notFound :        # Report this record as not found
                continue

            withoutAltUR += 1

            # Report the probable finds based upon Extensive checking
            if d.Extensive and (d.secondaryRecNo in d.possExtensiveFinds):
                f.PrintExtensiveFinds()

            # Report the findings based upon d.recStatus[d.secondarRecNo] for records that have been checked
            if d.secondaryRecNo not in d.recStatus:
                continue
            if d.recStatus[d.secondaryRecNo] == -1:
                f.PrintSecondary(True, '', 2, 'nf', f'not found in {d.masterLongName}', 0)
                notFoundtd += 1
            else:
                finds = 0    # Count multiple UR matches
                URs = {}
                d.URrecNos = []
                d.URrecSounds = []
                foundMasterRecNo = 0
                foundMasterSound = ''
                masterRecs = d.foundRec[d.secondaryRecNo].split('~')
                masterSounds = d.foundSound[d.secondaryRecNo].split('~')
                # we may have multiple findings - hopefully all the aliases and merged patients point back to the one "real" patient
                for i, masterRecNo in enumerate(masterRecs):
                    masterRecNo = int(masterRecNo)
                    masterRecSound = masterSounds[i]
                    thisUR = d.masterDetails[masterRecNo]['UR']
                    if masterRecNo in d.masterNewRec :            # merge
                        newMasterRecNo = d.masterNewRec[masterRecNo]
                        thisUR = d.masterDetails[newMasterRecNo]['UR']
                    elif masterRecNo in d.masterPrimRec :            # alias
                        newMasterRecNo = d.masterPrimRec[masterRecNo]
                        thisUR = d.masterDetails[newMasterRecNo]['UR']
                    if thisUR not in URs :            # Check repeat
                        URs[thisUR] = 1            # Save UR number
                        d.URrecNos.append(masterRecNo)
                        d.URrecSounds.append(masterRecSound)
                        foundMasterRecNo = masterRecNo
                        foundMasterSound = masterRecSound
                        finds += 1            # Count found
                if finds > 1 :                # Multiple finds / multiple patients
                    if d.recStatus[d.secondaryRecNo] == 6:
                        f.PrintDuplicateFound('')
                    elif d.recStatus[d.secondaryRecNo] == 5:
                        f.PrintDuplicateFound('Sound')
                    elif d.recStatus[d.secondaryRecNo] == 4:
                        f.PrintDuplicateFound('Sex')
                    elif d.recStatus[d.secondaryRecNo] == 3:
                        f.PrintDuplicateFound('Birthdate')
                    elif d.recStatus[d.secondaryRecNo] == 2:
                        f.PrintDuplicateFound('Given Name')
                    elif d.recStatus[d.secondaryRecNo] == 1:
                        f.PrintDuplicateFound('Surname')
                else :                    # Only one, it's an exact or partial match
                    isFound = f.CheckIfFound()
                    if d.recStatus[d.secondaryRecNo] == 6:
                        f.PrintMatchFound(foundMasterRecNo)
                        continue
                    if d.recStatus[d.secondaryRecNo] == 5:
                        f.PrintPartialFound('pfsnd', foundMasterRecNo, 'Sound', foundMasterSound)
                        pfoundsnd += 1
                        if isFound:
                            pfoundsnddn += 1
                    elif d.recStatus[d.secondaryRecNo] == 4:
                        f.PrintPartialFound('pfnsx', foundMasterRecNo, 'Sex', foundMasterSound)
                        pfoundsx += 1
                        if isFound:
                            pfoundsxdn += 1
                    elif d.recStatus[d.secondaryRecNo] == 3:
                        f.PrintPartialFound('pfnbd', foundMasterRecNo, 'Birthdate', foundMasterSound)
                        pfoundbd += 1
                        if isFound:
                            pfoundbddn += 1
                    elif d.recStatus[d.secondaryRecNo] == 2:
                        f.PrintPartialFound('pfng', foundMasterRecNo, 'Given Name', foundMasterSound)
                        pfoundg += 1
                        if isFound:
                            pfoundgdn += 1
                    elif d.recStatus[d.secondaryRecNo] == 1:
                        f.PrintPartialFound('pfnf', foundMasterRecNo, 'Family Name', foundMasterSound)
                        pfoundf += 1
                        if isFound:
                            pfoundfdn += 1


    if d.secondaryExtractDir:
        f.PrintClose('fdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Found_Done_{d.foundDoneVolume}.xlsx')
        f.PrintClose('ftd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Found_ToDo_{d.foundToDoVolume}.xlsx')
        f.PrintClose('nf', 0, 5, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_NotFound_ToDo.xlsx')
        f.PrintClose('fm', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_FoundMerged.xlsx')
        f.PrintClose('fa', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_FoundAlias.xlsx')
        f.PrintClose('dfdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateFound_Done.xlsx')
        f.PrintClose('dftd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateFound_ToDo.xlsx')
        f.PrintClose('dpfdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateSimilarFound_Done.xlsx')
        f.PrintClose('dpftd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateSimilarFound_ToDo.xlsx')
        f.PrintClose('pfsnddn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_OnSound_Done.xlsx')
        f.PrintClose('pfsndtd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_OnSound_ToDo.xlsx')
        f.PrintClose('pfnfdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotFamilyName_Done.xlsx')
        f.PrintClose('pfnftd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotFamilyName_ToDo.xlsx')
        f.PrintClose('pfnbddn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotBirthdate_Done.xlsx')
        f.PrintClose('pfnbdtd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotBirthdate_ToDo.xlsx')
        f.PrintClose('pfngdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotGivenName_Done.xlsx')
        f.PrintClose('pfngtd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotGivenName_ToDo.xlsx')
        f.PrintClose('pfnsxdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotSex_Done.xlsx')
        f.PrintClose('pfnsxtd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_SimilarFound_NotSex_ToDo.xlsx')
        f.PrintClose('ef', 0, 10, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Extensive_Finds.xlsx')
        f.PrintClose('pf', 0, 10, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Probable_Finds.xlsx')
    else:
        f.PrintClose('fdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_Found_Done_{d.foundDoneVolume}.xlsx')
        f.PrintClose('ftd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_Found_ToDo_{d.foundToDoVolume}.xlsx')
        f.PrintClose('nf', 0, 5, f'{d.secondaryDir}/{d.secondaryShortName}_NotFound_ToDo.xlsx')
        f.PrintClose('fm', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_FoundMerged.xlsx')
        f.PrintClose('fa', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_FoundAlias.xlsx')
        f.PrintClose('dfdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_DuplicateFound_Done.xlsx')
        f.PrintClose('dftd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_DuplicateFound_ToDo.xlsx')
        f.PrintClose('dpfdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_DuplicateSimilarFound_Done.xlsx')
        f.PrintClose('dpftd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_DuplicateSimilarFound_ToDo.xlsx')
        f.PrintClose('pfsnddn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_OnSound_Done.xlsx')
        f.PrintClose('pfsndtd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_OnSound_ToDo.xlsx')
        f.PrintClose('pfnfdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotFamilyName_Done.xlsx')
        f.PrintClose('pfnftd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotFamilyName_ToDo.xlsx')
        f.PrintClose('pfnbddn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotBirthdate_Done.xlsx')
        f.PrintClose('pfnbdtd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotBirthdate_ToDo.xlsx')
        f.PrintClose('pfngdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotGivenName_Done.xlsx')
        f.PrintClose('pfngtd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotGivenName_ToDo.xlsx')
        f.PrintClose('pfnsxdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotSex_Done.xlsx')
        f.PrintClose('pfnsxtd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_SimilarFound_NotSex_ToDo.xlsx')
        f.PrintClose('ef', 0, 10, f'{d.secondaryDir}/{d.secondaryShortName}_Extensive_Finds.xlsx')
        f.PrintClose('pf', 0, 10, f'{d.secondaryDir}/{d.secondaryShortName}_Probable_Finds.xlsx')

    logging.info('End of Pass 5')



    # Now printout the report
    f.openReport()
    heading = 'Phase 2 Testing - findUR'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')
    heading = f'{d.secondaryLongName} Patient Found in {d.masterLongName}'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')

    d.rpt.write(f'{count}\t{d.secondaryLongName} Patients Read In\n')
    d.rpt.write(f'{withoutAltUR}\t{d.secondaryLongName} Patients checked (patients without a {d.secondaryAltURname} number or without a valid {d.secondaryAltURname})\n\n')
    last = 0
    if (d.founddn + d.foundtd) > 0:
        d.rpt.write(f'{(d.founddn + d.foundtd)}\t{d.secondaryLongName} patients with only one similar {d.masterLongName} patient found\n')
        if d.founddn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{d.founddn} already checked - found.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Found_Done_\'n\'.xlsx)\n')
            else:
                d.rpt.write(f'\t\t[{d.founddn} already checked - found.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_Found_Done_n.xlsx)\n')
        if d.foundtd > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{d.foundtd} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Found_ToDo_\'n\'.xlsx)\n')
            else:
                d.rpt.write(f'\t\t[{d.foundtd} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_Found_ToDo_n.xlsx)\n')
        last += 1
    if d.mfound > 0:
        d.rpt.write(f'{d.mfound}\t{d.secondaryLongName} patients found to be similar to one {d.masterLongName} patient\n')
        d.rpt.write(f'\t\twho has been merged to a patient with a different {d.masterURname} ({d.secondaryShortName}_FoundMerged.xlsx)\n')
        last += 1
    if d.afound > 0:
        d.rpt.write(f'{d.afound}\t{d.secondaryLongName} patients found to be similar to one {d.masterLongName} patient\n')
        d.rpt.write(f'\t\twho is an alias of a patient with a differnt {d.masterURname} ({d.secondaryShortName}_FoundAlias.xlsx)\n')
        last += 1
    if (notFounddn + notFoundtd) > 0:
        d.rpt.write(f'{notFounddn + notFoundtd}\t{d.secondaryLongName} patients could not be found in {d.masterLongName}\n')
        if notFounddn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{notFounddn} already checked - {d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_NotFound_Done.xlsx]\n')
            else:
                d.rpt.write(f'\t\t[{notFounddn} already checked - {d.secondaryDir}/{d.secondaryShortName}_NotFound_Done.xlsx]\n')
        if notFoundtd > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{notFoundtd} to be checked - {d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_NotFound_ToDo.xlsx]\n')
            else:
                d.rpt.write(f'\t\t[{notFoundtd} to be checked - {d.secondaryDir}/{d.secondaryShortName}_NotFound_ToDo.xlsx]\n')
        last += 1
    if (d.dfounddn + d.dfoundtd) > 0:
        d.rpt.write(f'{d.dfounddn + d.dfoundtd}\t{d.secondaryLongName} patients with more than one similar {d.masterLongName} patient found\n')
        if d.dfounddn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{d.dfounddn} already checked - found.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateFound_Done.xlsx)\n')
            else:
                d.rpt.write(f'\t\t[{d.dfounddn} already checked - found.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_DuplicateFound_Done.xlsx)\n')
        if d.dfoundtd > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{d.dfoundtd} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateFound_ToDo.xlsx)\n')
            else:
                d.rpt.write(f'\t\t[{d.dfoundtd} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_DuplicateFound_ToDo.xlsx)\n')
        last += 1
    if (d.dpfounddn + d.dpfoundtd) > 0:
        d.rpt.write(f'{d.dpfounddn + d.dpfoundtd}\t{d.secondaryLongName} patients with more than one similar {d.masterLongName} patient found\n')
        if d.dpfounddn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{d.dpfounddn} already checked - found.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateSimilarFound_Done.xlsx)\n')
            else:
                d.rpt.write(f'\t\t[{d.dpfounddn} already checked - found.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_DuplicateSimilarFound_Done.xlsx)\n')
        if d.dpfoundtd > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t\t[{d.dpfoundtd} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_DuplicateSimilarFound_ToDo.xlsx)\n')
            else:
                d.rpt.write(f'\t\t[{d.dpfoundtd} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_DuplicateSimilarFound_ToDo.xlsx)\n')
        last += 1
    if pfoundf > 0:
        d.rpt.write(f'{pfoundf}\t{d.secondaryLongName} patients found to be similar to one {d.masterLongName} patient\n')
        d.rpt.write('\t\tmismatch on family [given name, birth date and sex match]\n')
        if pfoundfdn > 0:
            d.rpt.write(f'\t\t[{pfoundfdn} already checked - found.xlsx] ({d.secondaryShortname}_SimilarFound_NotFamilyName_Done.xlsx)\n')
        if (pfoundf - pfoundfdn) > 0:
            d.rpt.write(f'\t\t[{pfoundf - pfoundfdn} to be checked] ({d.secondaryShortName}_SimilarFound_NotFamilyName_ToDo.xlsx)\n')
        last += 1
    if pfoundg > 0:
        d.rpt.write(f'{pfoundg}\t{d.secondaryLongName} patients found to be simliar to one {d.masterLongName} patient\n')
        d.rpt.write('\t\tmismatch on given name [surname, birthdate and sex match]\n')
        if pfoundgdn > 0:
            d.rpt.write(f'\t\t[{pfoundgdn} already checked - found.xlsx] ({d.seconaryShortName}_SimilarFound_NotGivenName_Done.xlsx)\n')
        if (pfoundg - pfoundgdn) > 0:
            d.rpt.write(f'\t\t[{pfoundg - pfoundgdn} to be checked] ({d.secondaryShortName}_SimilarFound_NotGivenName_ToDo.xlsx)\n')
        last += 1
    if pfoundbd > 0:
        d.rpt.write(f'{pfoundbd}\t{d.secondaryLongName} patients found to be similar to one {d.masterLongName} patient\n')
        d.rpt.write('\t\tmismatch on birthdate [surname, firstname and sex match]\n')
        if pfoundbddn > 0:
            d.rpt.write(f'\t\t[{pfoundbddn} already checked - found.xlsx] ({d.secondaryShortName}_SimilarFound_NotBirthdate_Done.xlsx)\n')
        if (pfoundbd - pfoundbddn) > 0:
            d.rpt.write(f'\t\t[{pfoundbd - pfoundbddn} to be checked] ({d.secondaryShortName}_SimilarFound_NotBirthdate_ToDo.xlsx)\n')
        last += 1
    if pfoundsx > 0:
        d.rpt.write(f'{pfoundsx}\t{d.secondaryLongName} patients found to be similar to one {d.masterLongName} patient\n')
        d.rpt.write('\t\tmismatch on sex [surname, firstname and birthdate match]\n')
        if pfoundsxdn > 0:
            d.rpt.write(f'\t\t[{pfoundsxdn} already checked - found.xlsx] ({d.secondaryShortName}_SimilarFound_NotSex_Done.xlsx)\n')
        if (pfoundsx - pfoundsxdn) > 0:
            d.rpt.write(f'\t\t[{pfoundsx - pfoundsxdn} to be checked] ({d.secondaryShortName}_SimilarFound_NotSex_ToDo.xlsx)\n')
        last += 1
    if pfoundsnd > 0:
        d.rpt.write(f'{pfoundsnd}\t{d.secondaryLongName} patients found to similar to one {d.masterLongName} patient\n')
        d.rpt.write('\t\tmatched on birth date, sex, sound of surname and sound of firstname\n')
        if pfoundsnddn > 0:
            d.rpt.write(f'\t\t[{pfoundsnddn} already checked - found.xlsx] ({d.secondaryShortName}_SimilarFound_OnSound_Done.xlsx)\n')
        if (pfoundsnd - pfoundsnddn) > 0:
            d.rpt.write(f'\t\t[{pfoundsnd - pfoundsnddn} to be checked] ({d.secondaryShortName}_SimilarFound_OnSound_ToDo.xlsx)\n')
        last += 1
    d.rpt.write('\nNOTE: The ')
    if last == 1:
        d.rpt.write('last number should equal ')
    elif last == 2:
        d.rpt.write('last two numbers should add up to ')
    elif last == 3:
        d.rpt.write('last three numbers should add up to ')
    elif last == 4:
        d.rpt.write('last four numbers should add up to ')
    elif last == 5:
        d.rpt.write('last five numbers should add up to ')
    elif last == 6:
        d.rpt.write('last six numbers should add up to ')
    elif last == 7:
        d.rpt.write('last seven numbers should add up to ')
    elif last == 8:
        d.rpt.write('last eight numbers should add up to ')
    elif last == 9:
        d.rpt.write('last nine numbers should add up to ')
    elif last == 10:
        d.rpt.write('last ten numbers should add up to ')
    elif last == 11:
        d.rpt.write('last eleven numbers should add up to ')
    elif last == 12:
        d.rpt.write('last twelve numbers should add up to ')
    else:
        d.rpt.write('last thirteen numbers should add up to ')
    d.rpt.write('the second number\n')
    d.rpt.write('\n\n')


    # Report the possible matches based upon Extensive checking
    if d.Extensive:
        d.rpt.write(f'{d.extensivefind}\tRecords found to be identical, using extensive matching, to patients in {d.masterLongName} ({d.secondaryShortName}_Extensive_Finds.xlsx)\n')
        d.rpt.write(f'{d.probablefind}\tRecords found to be similar, using extensive matching, to patients in {d.masterLongName} ({d.secondaryShortName}_Probable_Finds.xlsx)\n')

    d.rpt.close()

    # Close the error log csv file and exit
    d.fe.close()
    sys.exit(0)
