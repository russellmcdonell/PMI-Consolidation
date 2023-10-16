
# pylint: disable=line-too-long

'''
A script to read in the secondary PMI file of patients with a alternate UR number.
Then check that the matching master PMI file record (same UR number) is a good match (or report how good a match it is).

SYNOPSIS
$ python matchAltUR.py masterDirectory [-e masterExtractDirectory|--masterExtractDir=masterExtractDirectory]
                    secondaryDirectory [-f secondaryExtractDirectory|--secondaryExtractDir=secondaryExtractDirectory]
                    [-E|--Extensive] [-S|--SkipMatched]
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

-S|--SkipMatched
Don't recheck things marked an matches (in matched.xlsx) or non-matches (in notMatched.xlsx)
This will suppress the creation of '_Done' files

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

Then read in the Secondary PMI file and create a "key" for each patient that has a alternate UR number.
(If the secondary PMI file contains deleted, merged or alias patients, then we ignore them)
That "key" will be made up of the cleaned up family name, cleaned up given name, birthdate and sex of the patient.

Next read in the master PMI file and compute the same "key" for each patient.
(We only do this if the Master PMI record has a UR number that has been found amoungst the alt UR numbers in the secondary PMI file)
We match the alt UR number in the secondary PMI file with the UR number in the Master PMI file.
Then we check to see how well the other fields match. To do this we check the "keys". If they match the it assumed to be a perfect match.
If not, we break the "keys" down into their component part and see how many of those components match.
If family name or given name aren't a perfect match, compute the sound confidence of a match
If the Extensive checking option is invoked, then compute and report an overall confidence score for best match based upon the "key".

The Master PMI file can contains aliases and merged patients. Look for both full and partial matches against aliases and
report the match against the alias as the alias and the primary patient are assumed to be a perfect match.
Determine which patients in the Master PMI file are records of interest. Then re-read the Master PMI file and save the information about patient of interest.

Finally, re-read the secondary PMI file and printout the matching information about patients of interest from the master PMI file

'''

# pylint: disable=invalid-name, bare-except, pointless-string-statement, unspecified-encoding

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
    d.scriptType = 'match'

    # Get the options
    parser = argparse.ArgumentParser(description='Look for matches between the secondary PMI file and the master PMI file, based upon AltUR === UR')
    parser.add_argument ('masterDir', metavar='masterDirectory', help='The name of directory containg the master configuration and master PMI file (master.csv)')
    parser.add_argument ('secondaryDir', metavar='secondaryDirectory', help='The name of directory containg the secondary configuration secondary PMI file (secondary.csv)')
    parser.add_argument ('-e', '--masterExtractDir', dest='masterExtractDir', metavar='masterExtractDirectory', help='The name of the master directory sub-directory that contains the extract master CSV file and configuration specific to the extract')
    parser.add_argument ('-f', '--secondaryExtractDir', dest='secondaryExtractDir', metavar='secondaryExtractDirectory', help='The name of the secondary directory sub-directory that contains the extract secondary CSV file and configuration specific to the extract')
    parser.add_argument ('-E', '--Extensive', dest='Extensive', action='store_true', help='Invoke Extensive checking to look for possible matches')
    parser.add_argument ('-S', '--SkipMatched', dest='skipMatched', action='store_true', help='Skip testing matches/non-matches (suppress _Done files)')
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
        if args.logfile:        # and send it to a file if the -o logfile option is specified
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', level=logging_levels[loggingLevel], filename=args.logfile)
        else:
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', level=logging_levels[loggingLevel])
    else:
        if args.logfile:        # send the default (WARN) logging to a file if the -o logfile option is specified
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p', filename=args.logfile)
        else:
            logging.basicConfig(format=logfmt, datefmt='%d/%m/%y %H:%M:%S %p')


    d.masterDir = args.masterDir
    d.secondaryDir = args.secondaryDir
    d.masterExtractDir = args.masterExtractDir
    d.secondaryExtractDir = args.secondaryExtractDir
    d.Extensive = args.Extensive
    skipMatched = args.skipMatched
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
        sys.exit(EX_SOFTWARE)
    try:
        loader = SourceFileLoader('mc', f'./{d.masterDir}/cleanMaster.py')
        spec = spec_from_loader('mc', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.mc = module
    except:
        logging.fatal('importing ./%s/cleanMaster.py failed', d.masterDir)
        sys.exit(EX_SOFTWARE)
    if not os.path.exists(f'./{d.masterDir}/linkMaster.py'):
        logging.fatal('./%s/linkMaster.py not found', d.masterDir)
        sys.exit(EX_SOFTWARE)
    try:
        loader = SourceFileLoader('ml', f'./{d.masterDir}/linkMaster.py')
        spec = spec_from_loader('ml', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.ml = module
    except:
        logging.fatal('importing ./%s/linkMaster.py failed', d.masterDir)
        sys.exit(EX_SOFTWARE)

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
        sys.exit(EX_SOFTWARE)
    try:
        loader = SourceFileLoader('sc', f'./{d.secondaryDir}/cleanSecondary.py')
        spec = spec_from_loader('sc', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.sc = module
    except:
        logging.fatal('importing ./%s/cleanSecondary.py failed', d.secondaryDir)
        sys.exit(EX_SOFTWARE)
    if not os.path.exists(f'./{d.secondaryDir}/linkSecondary.py'):
        logging.fatal('./%s/linkSecondary.py not found', d.secondaryDir)
        sys.exit(EX_SOFTWARE)
    try:
        loader = SourceFileLoader('sl', f'./{d.secondaryDir}/linkSecondary.py')
        spec = spec_from_loader('sl', loader)
        module = module_from_spec(spec)
        loader.exec_module(module)
        d.sl = module
    except:
        logging.fatal('importing ./%s/linkSecondary.py failed', d.secondaryDir)
        sys.exit(EX_SOFTWARE)

    # Open Error file
    f.openErrorFile()

    # Read in the Matched PIDs - matches found and no checking required
    f.getMatched()

    # Read in the Not Matched PIDs - confirmed non-matches - further checking requried
    f.getNotMatched()


    # Pass 1 - pick out the keys from the secondary PMI file for records that can be matched [those with an AltUR which isn't know to be bad]
    # Open the cleaned up CSV secondary PMI file
    secondaryCSV = None
    d.altUR = {}
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

            # Only check patients with an altUR number that isn't known to be wrong (notMatch.xlsx)
            altUR = d.sc.secondaryCleanAltUR()    # Get the altUR number
            if altUR == '':
                continue

            pid = d.sc.secondaryCleanPID()        # Get the PID number
            if pid in d.notMatched :        # Skip this record if the AltUR in this record is known to be wrong - use findUR to find a match for this record
                continue

            # Check if this pid has already been Matched (i.e. the pid and altUR are in matched.xlsx because the record has been previously matched)
            if skipMatched and (pid in d.matchedUR):
                # Check that the altUR matches the "known" altUR
                if altUR != d.matchedUR[pid]:
                    d.feCSV.writerow([f'{d.progName}:ERROR in matched.xlsx:{pid},{d.matchedUR[pid]} - Does not match {d.secondaryAltURname} for {d.secondaryLongName} {d.secondaryPIDname} {pid},{altUR} - ignoring this match'])
                else:
                    d.foundSecondaryRec[pid] = d.secondaryRecNo
                    d.secondaryFoundPID[d.secondaryRecNo] = pid
                    d.recStatus[d.secondaryRecNo] = 64
                    d.foundRec[d.secondaryRecNo] = -1        # The matching master PMI record for this secondary PMI record (-1 == none found yet/to be assigned)
                    d.extras[d.secondaryRecNo] = ''
                    continue

            # Save the record number(s) for this altUR
            # There should be only one, but we have to allow for multiples. Later on we go through any 'list' and look for the best match.
            # NOTE: the 'list' can contain merged and deleted records, but we check everything.
            if altUR not in d.altUR:
                d.altUR[altUR] = []
                d.altURrec[altUR] = None
            d.altUR[altUR].append(d.secondaryRecNo)

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
            d.fullKey[d.secondaryRecNo] = thisKey
            d.recStatus[d.secondaryRecNo] = -1
            d.foundRec[d.secondaryRecNo] = -1
            d.extras[d.secondaryRecNo] = ''

            if d.Extensive and pid not in d.foundSecondaryRec :        # Collect and pack Extensive checking data (if required)
                d.ExtensiveSecondaryRecKey[d.secondaryRecNo] = f.Sounds(Sf, Sg)
                if Sdob == '':
                    d.ExtensiveSecondaryBirthdate[d.secondaryRecNo] = d.futureBirthdate
                else:
                    d.ExtensiveSecondaryBirthdate[d.secondaryRecNo] = f.Birthdate(Sdob)
                if d.useMiddleNames:
                    d.ExtensiveSecondaryMiddleNames[d.secondaryRecNo] = f.secondaryField('MiddleNames').upper()
                thisHash = {}
                for field in (sorted(d.ExtensiveFields.keys())):
                    fieldData =  f.secondaryField(field)
                    if fieldData == '':
                        thisHash[field] = None
                    else:
                        thisHash[field] = hash(fieldData)
                d.ExtensiveOtherSecondaryFields[d.secondaryRecNo] = thisHash
    logging.info('End of Pass 1')


    # Pass 2 - read the Master PMI file and match keys updating status of records in the Secondary PMI file as you go
    #       Also, save the  merged patients links as you will need them later
    #       And build up the data about probabalistic matched (Possible Matches)
    # Open the cleaned up CSV master PMI file
    d.possExtensiveMatches = {}                # all the matches for each of the matched secondary PMI record (secondary AltUR matches master UR)
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

            # Check for Matched patients
            ur = d.mc.masterCleanUR()        # The UR number

            # Check for secondary PIDs that have already been matched to this UR (this UR is an AltUR in matched.xlsx)
            if (not d.ml.masterIsAlias()) and (not d.ml.masterIsMerged()) and (ur in d.matchedPID):
                secondaryPIDs = d.matchedPID[ur].split('~')            # The list of secondary PIDs with AltURs that match this master UR
                for secondaryPID in (secondaryPIDs) :                # Check that each of these secondary PIDs was found in the secondary PMI (see above)
                    if secondaryPID in d.foundSecondaryRec:
                        d.foundRec[d.foundSecondaryRec[secondaryPID]] = d.masterRecNo        # Save the master records number for this secondary record number
                    else:
                        d.feCSV.writerow(['{d.progName}:ERROR in matched.xlsx:{ur},{d.matchedPID[ur]} - {d.secondaryLongName} {d.secondaryPIDname} {secondaryPID} not found'])


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

            # Check if we are interested in this Master PMI UR (i.e. not an AltUR found in the secondary PMI file)
            if ur not in d.altUR:
                if testKey:
                    logging.info('\t NOT CHECKED - %s %s not an %s in the secondary PMI file', d.masterURname, ur, d.secondaryAltURname)
                continue

            # Collect Extensive checking data if required
            Mfny = Mfdm =  Mfsx = Mgny = Mgdm = Mgsx = My = Mm = Md = masterBirthdate = Mmn = None
            masterOtherFields = {}
            if d.Extensive:
                sounds = f.Sounds(Mf, Mg)
                (Mfny, Mfdm, Mfsx, Mgny, Mgdm, Mgsx) = sounds.split('~')
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

            # Check this Master PMI file record against each secondary PMI record which has an altUR matching this UR
            bestMatch = -1
            for secRecNo in (d.altUR[ur]) :            # Each secondary record which has this UR as it's AltUR
                if testKey:
                    logging.info('\t0:%s:%s:%s:%s:%s', secRecNo, d.fullKey[secRecNo], d.recStatus[secRecNo], d.foundRec[secRecNo], d.extras[secRecNo])

                # Extact the key componenents
                (Sf, Sg, Sdob, Ssex) = d.fullKey[secRecNo].split('~')
                Sfny = Sfdm =  Sfsx = Sgny = Sgdm = Sgsx = None

                # For Extensive checking we compute a confidence level that this master record matches each of the associated secondary record
                # There can be multiple secondary records claiming to be linked to each master record
                if d.Extensive:
                    # Unpack the secondary PMI items for this secondary PMI record
                    (Sfny, Sfdm, Sfsx, Sgny, Sgdm, Sgsx) = d.ExtensiveSecondaryRecKey[secRecNo].split('~')
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
                    for field, weight in d.ExtensiveFields.items():
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
                        if secRecNo not in d.possExtensiveMatches:
                            d.possExtensiveMatches[secRecNo] = {}
                        if totalConfidence not in (d.possExtensiveMatches[secRecNo]):
                            d.possExtensiveMatches[secRecNo][totalConfidence] = []
                        d.possExtensiveMatches[secRecNo][totalConfidence].append([d.masterRecNo, soundFamilyNameConfidence, soundGivenNameConfidence])


                # If forced, or a perfect match has already been found, for the secondary PMI file record then don't do any further checking
                if d.recStatus[secRecNo] > 62:
                    if d.recStatus[secRecNo] > bestMatch:
                        d.altURrec[ur] = secRecNo
                        d.foundRec[secRecNo] = d.masterRecNo
                        bestMatch = d.recStatus[secRecNo]
                    continue

                # Check for a perfect match
                if d.fullKey[secRecNo] == thisKey:
                    d.recStatus[secRecNo] = 63
                    d.foundRec[secRecNo] = d.masterRecNo
                    d.extras[secRecNo] = ''
                    if d.recStatus[secRecNo] > bestMatch:
                        d.altURrec[ur] = secRecNo
                        bestMatch = d.recStatus[secRecNo]
                    if testKey:
                        logging.info('\t1:%s:%s:%s:%s:%s', secRecNo, d.fullKey[secRecNo], d.recStatus[secRecNo], d.foundRec[secRecNo], d.extras[secRecNo])
                    continue
                else:
                    if testKey:
                        logging.info('\tNOT MATCHED!!!')

                # Check for one of the partial matches
                thisStatus = 0
                extras = ''        # Record extra things that didn't match
                if Sf != Mf :            # Check Family Names
                    # If a better match for this record has already been found then skip this record
                    if d.recStatus[secRecNo] > 31 :        # max score without family name match (family name sound(8) + dob(16) + given name(4+2) + sex(1))
                        continue

                    (fnMatch, weight) = f.FamilyNameSoundCheck(Mf, Mfny, Mfdm, Mfsx, Sf, Sfny, Sfdm, Sfsx, 1.0)
                    if fnMatch > 0.7:
                        thisStatus = 8
                        extras = f'(FN sound [{fnMatch:0.2f}])'
                else:
                    thisStatus = 32 + 8

                if Sdob != Mdob :        # Next check Date of Birth - may add to matching score
                    # If a better match for this record has already been found then skip this record
                    if d.recStatus[secRecNo] > thisStatus + 23 :        # max score with dob(16) + given name(4+2) + sex(1)
                        continue

                    if (thisStatus & 32) == 0:
                        if extras != '':
                            extras += ', '
                        extras += 'plus birthdate'
                else:
                    thisStatus += 16

                if Sg != Mg :        # Next check Given Name - may add to matching score
                    # If a better match for this record has already been found then skip this record
                    if d.recStatus[secRecNo] > thisStatus + 3 :        # max score without given name match (given name sound(2) + sex(1))
                        continue

                    if (thisStatus & 48) == 0:
                        if extras != '':
                            extras += ', '
                        extras += 'plus given name'
                    (gnMatch, weight) = f.GivenNameSoundCheck(Mg, Mgny, Mgdm, Mgsx, Sg, Sgny, Sgdm, Sgsx, 1.0)
                    if gnMatch > 0.7:
                        thisStatus += 2
                        if extras != '':
                            extras += ', '
                        extras += f'(GN sound [{gnMatch:0.2f}])'
                else:
                    thisStatus += 4 + 2

                if Ssex != Msex :        # Next check Sex - may add to matching score
                    # If a better match for this record has already been found then skip this record
                    if d.recStatus[secRecNo] > thisStatus :        # max score with sex
                        continue

                    if (thisStatus & 52) == 0:
                        if extras != '':
                            extras += ', '
                        extras += 'plus sex'
                else:
                    thisStatus += 1

                # And no better match has already been found
                if thisStatus == d.recStatus[secRecNo] :            # An equal match has already been found - update if this it the primary record
                    if d.ml.masterIsAlias():
                        continue
                    if d.ml.masterIsMerged() and (d.masterMerged != 'IN'):
                        continue
                d.extras[secRecNo] = extras
                d.recStatus[secRecNo] = thisStatus
                d.foundRec[secRecNo] = d.masterRecNo
                if d.recStatus[secRecNo] > bestMatch:
                    d.altURrec[ur] = secRecNo
                    bestMatch = d.recStatus[secRecNo]
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
            masterRecNo = d.foundRec[secRecNo]
            if masterRecNo == -1:
                d.feCSV.writerow([f'{d.progName}:ERROR in matched.xlsx:{d.secondaryFoundPID[secRecNo]},{d.matchedUR[d.secondaryFoundPID[secRecNo]]} - {d.masterLongName} {d.masterURname} number not found'])
            else:
                # Mark as wanted
                d.wantedMasterRec[masterRecNo] = True
                # Check if an alias or merged patient and if so get master record as well
                if masterRecNo in d.masterPrimRec:
                    d.wantedMasterRec[d.masterPrimRec[masterRecNo]] = True
                if masterRecNo in d.masterNewRec:
                    d.wantedMasterRec[d.masterNewRec[masterRecNo]] = True

    for secRecNo, confidences in d.possExtensiveMatches.items():
        for confidence, recNumbers in confidences.items():
            for i, thisRecNo in enumerate(recNumbers):
                masterRecNo = thisRecNo[0]
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
    d.matchVolume = 1
    d.date_style = NamedStyle(name='Date', number_format='dd-mmm-yyyy')
    for thisFile in ['fnc', 'fno', 'fnp', 'gn', 'db', 'sex']:
        f.PrintHeading(thisFile + 'td', 3, 'Checked', 'Message')
        f.PrintHeading(thisFile + 'dn', 3, 'Checked', 'Message')
    if not d.quick:
        f.PrintHeading('m', 3, 'Message', '')
    f.PrintHeading('mm', 3, 'Message', '')
    f.PrintHeading('ma', 3, 'Message', '')
    f.PrintHeading('ud', 2, 'Message', '')
    if d.Extensive:
        f.PrintHeading('em', 4, 'Extensive Match', '')
        f.PrintHeading('pm', 4, 'Probable Match', '')

    # Open the cleaned up CSV secondary PMI file
    withoutAltUR = 0    # Count secondary records without altUR
    withBadAltUR = 0    # Count secondary records with bad altUR [matchNotMatched.csv]
    withAltUR = 0        # Count secondary records with altUR numbers
    urundef = 0        # Count of secondary records with unmatched altUR number
    d.mmatch = 0        # Count of secondary records matched to a merged master PMI patent
    d.amatch = 0        # Count of secondary records matched to an alias master PMI patent
    d.match = 0        # Count of secondary records matched to only one master PMI patent
    d.fncmis = 0        # Count of secondary records matched, but family not matched - close
    d.fncmisdn = 0        # Count of secondary records matched, but family not matched - close (Done)
    d.fnomis = 0        # Count of secondary records matched, but family not matched - only
    d.fnomisdn = 0        # Count of secondary records matched, but family not matched - only (Done)
    d.fnpmis = 0        # Count of secondary records matched, but family not matched - plus
    d.fnpmisdn = 0        # Count of secondary records matched, but family not matched - plus (Done)
    d.dobmis = 0        # Count of secondary records matched, but birthdate did not match
    d.dobmisdn = 0        # Count of secondary records matched, but birthdate did not match (Done)
    d.gnmis = 0        # Count of secondary records matched, but given name did not match
    d.gnmisdn = 0        # Count of secondary records matched, but given name did not match (Done)
    d.sexmis = 0        # Count of secondary records matched, but sex did not match
    d.sexmisdn = 0        # Count of secondary records matched, but sex did not match (Done)
    d.probmatch = 0        # Count of secondary records matched as probable matches using extensive matching
    d.extensivematch = 0    # Count of secondary records matched exactly using extensive matching
    d.probablematch = 0    # Count of secondary records matched approximately using extensive matching

    # Re-read the secondary PMI file
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

            # Only check patients with an altUR number
            altUR = d.sc.secondaryCleanAltUR()    # Get the altUR number
            if altUR == '':
                withoutAltUR += 1
                continue

            pid = d.sc.secondaryCleanPID()        # Get the PID number
            if pid in d.notMatched :        # Skip this record if the AltUR is known to be wrong - have to use findUR for this record
                withBadAltUR += 1
                continue

            withAltUR += 1

            # Report the possible matches based upon Extensive checking
            if d.Extensive and (d.secondaryRecNo in d.possExtensiveMatches):
                f.PrintExtensiveMatch()

            # Report the findings based upon d.recStatus[d.secondarRecNo]
            masterRecNo = d.foundRec[d.secondaryRecNo]
            if d.recStatus[d.secondaryRecNo] == -1:
                f.PrintNoMatch()
                urundef += 1
            elif d.recStatus[d.secondaryRecNo] >= 63:
                f.PrintMatchFound(masterRecNo)
            elif (d.recStatus[d.secondaryRecNo] & 32) == 0:
                if (d.recStatus[d.secondaryRecNo] & 8) != 0:
                    f.PrintMisMatch(masterRecNo, 'fnc', 'Family Name [close]')
                elif (d.recStatus[d.secondaryRecNo] & 23) == 23:
                    f.PrintMisMatch(masterRecNo, 'fno', 'Family Name Only')
                else:
                    f.PrintMisMatch(masterRecNo, 'fnp', 'Family Name Plus')
            elif (d.recStatus[d.secondaryRecNo] & 16) == 0:
                f.PrintMisMatch(masterRecNo, 'db', 'Birthdate')
            elif (d.recStatus[d.secondaryRecNo] & 4) == 0:
                f.PrintMisMatch(masterRecNo, 'gn', 'Given Names')
            else:
                f.PrintMisMatch(masterRecNo, 'sex', 'Sex')
    logging.info('End of Pass 5')


    # And finally, create the report
    f.openReport()
    heading = 'Phase 1 Testing - matchAltUR'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')
    heading = f'{d.secondaryLongName} Patient Matched on {d.secondaryAltURname} number'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')

    d.rpt.write(f'{withAltUR}\twith an {d.secondaryAltURname} number\n')
    d.rpt.write(f'{withoutAltUR}\twithout an {d.secondaryAltURname} number\n')
    if withBadAltUR > 0:
        d.rpt.write(f'{withBadAltUR}\twith a bad {d.secondaryAltURname} number\n')
    d.rpt.write(f'{withAltUR + withoutAltUR + withBadAltUR}\tTotal Patients\n\n')
    lastCount = 0
    if d.match > 0:
        if d.secondaryExtractDir:
            d.rpt.write(f'{d.match}\tRecords perfectly matched patients in {d.masterLongName} ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_matched_\'n\'.xlsx)\n')
            if not d.quick:
                f.PrintClose('m', 0, 6, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_matched_{d.matchVolume}.xlsx')
        else:
            d.rpt.write(f'{d.match}\tRecords perfectly matched patients in {d.masterLongName} ({d.secondaryDir}/{d.secondaryShortName}_ur_matched_\'n\'.xlsx)\n')
            if not d.quick:
                f.PrintClose('m', 0, 6, f'{d.secondaryDir}/{d.secondaryShortName}_ur_matched_{d.matchVolume}.xlsx')
        lastCount += 1
    if d.mmatch > 0:
        if d.secondaryExtractDir:
            d.rpt.write(f'{d.mmatch}\tRecords matched patients who have been merged to a patient with a different {d.masterURname} in {d.masterLongName} ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_matched_merged.xlsx)\n')
            f.PrintClose('mm', 0, 6, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_matched_merged.xlsx')
        else:
            d.rpt.write(f'{d.mmatch}\tRecords matched patients who have been merged to a patient with a different {d.masterURname} in {d.masterLongName} ({d.secondaryDir}/{d.secondaryShortName}_ur_matched_merged.xlsx)\n')
            f.PrintClose('mm', 0, 6, f'{d.secondaryDir}/{d.secondaryShortName}_ur_matched_merged.xlsx')
        lastCount += 1
    if d.amatch > 0:
        if d.secondaryExtractDir:
            d.rpt.write(f'{d.amatch}\tRecords matched patients who have an alias with a different {d.masterURname} in {d.masterLongName} ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_matched_merged.xlsx)\n')
            f.PrintClose('ma', 0, 6, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_matched_alias.xlsx')
        else:
            d.rpt.write(f'{d.amatch}\tRecords matched patients who have an alias with a different {d.masterURname} in {d.masterLongName} ({d.secondaryDir}/{d.secondaryShortName}_ur_matched_merged.xlsx)\n')
            f.PrintClose('ma', 0, 6, f'{d.secondaryDir}/{d.secondaryShortName}_ur_matched_alias.xlsx')
        lastCount += 1
    if urundef > 0:
        if d.secondaryExtractDir:
            d.rpt.write(f'{urundef}\tpatients had a {d.secondaryAltURname} number not found in {d.masterLongName} ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_undefined.xlsx)\n')
            f.PrintClose('ud', 0, 4, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ur_undefined.xlsx')
        else:
            d.rpt.write(f'{urundef}\tpatients had a {d.secondaryAltURname} number not found in {d.masterLongName} ({d.secondaryDir}/{d.secondaryShortName}_ur_undefined.xlsx)\n')
            f.PrintClose('ud', 0, 4, f'{d.secondaryDir}/{d.secondaryShortName}_ur_undefined.xlsx')
        lastCount += 1
    if d.fncmis > 0:
        d.rpt.write(f'{d.fncmis}\tRecords matched on {d.secondaryAltURname}/{d.masterURname} number, but not family name[family name sounds like matched]\n')
        if d.fncmisdn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.fncmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_close_mismatch_Done.xlsx)\n')
                f.PrintClose('fncdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_close_mismatch_Done.xlsx')
            else:
                d.rpt.write(f'\t[{d.fncmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_family_name_close_mismatch_Done.xlsx)\n')
                f.PrintClose('fncdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_family_name_close_mismatch_Done.xlsx')
        if (d.fncmis - d.fncmisdn) > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.fncmis - d.fncmisdn} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_close_mismatch_ToDo.xlsx)\n')
                f.PrintClose('fnctd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_close_mismatch_ToDo.xlsx')
            else:
                d.rpt.write(f'\t[{d.fncmis - d.fncmisdn} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_family_name_close_mismatch_ToDo.xlsx)\n')
                f.PrintClose('fnctd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_family_name_close_mismatch_ToDo.xlsx')
        lastCount += 1
    if d.fnomis > 0:
        d.rpt.write(f'{d.fnomis}\tRecords matched on {d.secondaryAltURname}/{d.masterURname} number, but not family name[given name, birth date and sex matched]\n')
        if d.fnomisdn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.fnomisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_only_mismatch_Done.xlsx)\n')
                f.PrintClose('fnodn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_only_mismatch_Done.xlsx')
            else:
                d.rpt.write(f'\t[{d.fnomisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_family_name_only_mismatch_Done.xlsx)\n')
                f.PrintClose('fnodn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_family_name_only_mismatch_Done.xlsx')
        if (d.fnomis - d.fnomisdn) > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.fnomis - d.fnomisdn} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_only_mismatch_ToDo.xlsx)\n')
                f.PrintClose('fnotd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_only_mismatch_ToDo.xlsx')
            else:
                d.rpt.write(f'\t[{d.fnomis - d.fnomisdn} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_family_name_only_mismatch_ToDo.xlsx)\n')
                f.PrintClose('fnotd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_family_name_only_mismatch_ToDo.xlsx')
        lastCount += 1
    if d.fnpmis > 0:
        d.rpt.write(f'{d.fnpmis}\tRecords matched on {d.secondaryAltURname}/{d.masterURname} number, but not family name[and at least one of given name, birth date or sex]\n')
        if d.fnpmisdn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.fnpmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_plus_mismatch_Done.xlsx)\n')
                f.PrintClose('fnpdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_plus_mismatch_Done.xlsx')
            else:
                d.rpt.write(f'\t[{d.fnpmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_family_name_plus_mismatch_Done.xlsx)\n')
                f.PrintClose('fnpdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_family_name_plus_mismatch_Done.xlsx')
        if (d.fnpmis - d.fnpmisdn) > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.fnpmis - d.fnpmisdn} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_plus_mismatch_ToDo.xlsx)\n')
                f.PrintClose('fnptd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_family_name_plus_mismatch_ToDo.xlsx')
            else:
                d.rpt.write(f'\t[{d.fnpmis - d.fnpmisdn} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_family_name_plus_mismatch_ToDo.xlsx)\n')
                f.PrintClose('fnptd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_family_name_plus_mismatch_ToDo.xlsx')
        lastCount += 1
    if d.dobmis > 0:
        d.rpt.write(f'{d.dobmis}\tRecords matched on {d.secondaryAltURname}/{d.masterURname} number and family name, but not birth date\n')
        if d.dobmisdn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.dobmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_dob_mismatch_Done.xlsx)\n')
                f.PrintClose('dbdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_dob_mismatch_Done.xlsx')
            else:
                d.rpt.write(f'\t[{d.dobmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_dob_mismatch_Done.xlsx)\n')
                f.PrintClose('dbdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_dob_mismatch_Done.xlsx')
        if (d.dobmis - d.dobmisdn) > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.dobmis - d.dobmisdn} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_dob_mismatch_ToDo.xlsx)\n')
                f.PrintClose('dbtd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_dob_mismatch_ToDo.xlsx')
            else:
                d.rpt.write(f'\t[{d.dobmis - d.dobmisdn} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_dob_mismatch_ToDo.xlsx)\n')
                f.PrintClose('dbtd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_dob_mismatch_ToDo.xlsx')
        lastCount += 1
    if d.gnmis > 0:
        d.rpt.write(f'{d.gnmis}\tRecords matched on {d.secondaryAltURname}/{d.masterURname} number, family name and birth date, but not given name\n')
        if d.gnmisdn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.gnmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_given_name_mismatch_Done.xlsx)\n')
                f.PrintClose('gndn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_given_names_mismatch_Done.xlsx')
            else:
                d.rpt.write(f'\t[{d.gnmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_given_name_mismatch_Done.xlsx)\n')
                f.PrintClose('gndn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_given_names_mismatch_Done.xlsx')
        if (d.gnmis - d.gnmisdn) > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.gnmis - d.gnmisdn} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_given_name_mismatch_ToDo.xlsx)\n')
                f.PrintClose('gntd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_given_names_mismatch_ToDo.xlsx')
            else:
                d.rpt.write(f'\t[{d.gnmis - d.gnmisdn} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_given_name_mismatch_ToDo.xlsx)\n')
                f.PrintClose('gntd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_given_names_mismatch_ToDo.xlsx')
        lastCount += 1
    if d.sexmis > 0:
        d.rpt.write(f'{d.sexmis}\tRecords matched on {d.secondaryAltURname}/{d.masterURname} number, family name, birth date and given name, but not sex\n')
        if d.sexmisdn > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.sexmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_sex_mismatch_Done.xlsx)\n')
                f.PrintClose('sexdn', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_sex_mismatch_Done.xlsx')
            else:
                d.rpt.write(f'\t[{d.sexmisdn} already checked - matched.xlsx] ({d.secondaryDir}/{d.secondaryShortName}_sex_mismatch_Done.xlsx)\n')
                f.PrintClose('sexdn', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_sex_mismatch_Done.xlsx')
        if (d.sexmis - d.sexmisdn) > 0:
            if d.secondaryExtractDir:
                d.rpt.write(f'\t[{d.sexmis - d.sexmisdn} to be checked] ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_sex_mismatch_ToDo.xlsx)\n')
                f.PrintClose('sextd', 0, 7, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_sex_mismatch_ToDo.xlsx')
            else:
                d.rpt.write(f'\t[{d.sexmis - d.sexmisdn} to be checked] ({d.secondaryDir}/{d.secondaryShortName}_sex_mismatch_ToDo.xlsx)\n')
                f.PrintClose('sextd', 0, 7, f'{d.secondaryDir}/{d.secondaryShortName}_sex_mismatch_ToDo.xlsx')
        lastCount += 1
    lastText = ['', '', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve']
    sumText = '\nNOTE: The last '
    if lastCount == 1:
        sumText += 'number'
    else:
        sumText += f'{lastText[lastCount]} numbers'
    sumText += ' should '
    if lastCount == 1:
        sumText += 'equal'
    else:
        sumText += 'add up to'
    sumText += ' the first\n'
    d.rpt.write(sumText)
    d.rpt.write('\n')

    # Report the possible matches based upon Extensive checking
    if d.Extensive:
        if d.secondaryExtractDir:
            d.rpt.write(f'{d.extensivematch}\tRecords exactly matched, using extensive matching, to patients in {d.masterLongName} ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Extensive_Matches.xlsx)\n')
            f.PrintClose('em', 1, 9, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Extensive_Matches.xlsx')
            d.rpt.write(f'{d.probablematch}\tRecords probably matched, using extensive matching, to patients in {d.masterLongName} ({d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Probable_Matches.xlsx)\n')
            f.PrintClose('pm', 1, 9, f'{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_Probable_Matches.xlsx')
        else:
            d.rpt.write(f'{d.extensivematch}\tRecords exactly matched, using extensive matching, to patients in {d.masterLongName} ({d.secondaryDir}/{d.secondaryShortName}_Extensive_Matches.xlsx)\n')
            f.PrintClose('em', 1, 9, f'{d.secondaryDir}/{d.secondaryShortName}_Extensive_Matches.xlsx')
            d.rpt.write(f'{d.probablematch}\tRecords probably matched, using extensive matching, to patients in {d.masterLongName} ({d.secondaryDir}/{d.secondaryShortName}_Probable_Matches.xlsx)\n')
            f.PrintClose('pm', 1, 9, f'{d.secondaryDir}/{d.secondaryShortName}_Probable_Matches.xlsx')


    d.rpt.close()

    # Close the error log csv file and exit
    d.fe.close()
    sys.exit(EX_OK)
