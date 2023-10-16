
# pylint: disable=line-too-long

'''
A script to check the goodness of health of a master PMI file

SYNOPSIS
$ python checkMaster.py masterDirectory [-e masterExtractDirectory|--masterExtractDir=masterExtractDirectory] [-E|--Extensive] [-m masterDebugKey|--masterDebugKey=masterDebugKey] [-n masterDebugCount|--masterDebugCount=masterDebugCount] [-q|--quick] [-v loggingLevel|--verbose=loggingLeve] [-o logfile|--logfile=logfile]


OPTIONS
masterDirectory
The directory containing the master configuration (master.cfg) plus the cleanMaster.py and linkMaster.py routines. Default is 'Master'
This directory may contain subdirectories where specific extracts will be found.

-e masterExtractDir|--masterExtractDir=masterExtractDir
The optional extract sub-directory, of the master directory, containing the extract master CSV file and parameters.

-E|--Extensive
Invoke extensive checking for possible duplicates. Extensive checking can invoke various function on the core data elements (FamilyName, GivenName, Birthdate, Sex)
plus simple checks of equality for any other field in the master PMI extract file.

-m masterDebugKey|--masterDebugKey=masterDebugKey
The key for triggering logging of information about a specific master record. Default is None

-n masterDebugCount|--masterDebugCount=masterDebugCount
A counter to trigger progress logging; a progress message is created every masterDebugCount(th) master record. Default is 50000

-q|--quick
Just performa a basic check of the master CSV file. Do not check alias or meged links. Do not create the cleaned up master CSV file.

-v loggingLevel|--verbose=loggingLevel
Set the level of logging that you want.

-o logfile|--logfile=logfile
The name of a log file where you want all messages captured.


THE MAIN CODE
Start by parsing the command line arguements and setting up logging.

Then read in the master PMI extract file and save it as a CSV file and make sure that there are no errors in the master PMI extract.
Many of the checks are contained in masterDirectory/cleanMaster.py
masterDirectory/cleanMaster.py may need editing to match the constraints for valid master records
Checks include records with the wrong number of fields, the UR number is valid, alias records have a matching master record,
merged records have a matching master record and probable duplicates.
This is based up cleaned up family name and cleaned up given name so there is no guaranttees that they are duplicates.
No checks of address, medicare number, next of kin or any other secondary identifiers are attempted unless the Extensive option is invoked
If the Extensive options is involked then a further check is conducted for possible duplicates.

The function in masterDirectory/cleanMaster.py associated with cleaning up family names, given names, birthdates and gender
may need editing to reflect any specific anomolies in the PMI extract.
'''

# pylint: disable=invalid-name, bare-except, pointless-string-statement, unspecified-encoding

# Import the required modules
import os
import sys
import csv
import argparse
import logging
import re
import datetime
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
Then read in the master file and save it as a CSV file, making sure that there are no errors in the master PMI extract.
    '''

    # Save the program name
    d.progName = sys.argv[0]
    d.progName = d.progName[0:-3]        # Strip off the .py ending
    d.scriptType = 'master'

    # Get the options
    parser = argparse.ArgumentParser(description='Check the goodness of health of a master PMI extraction')
    parser.add_argument ('masterDir', metavar='masterDirectory', help='The name of directory containg the master configuration and cleanMaster.py routines')
    parser.add_argument ('-e', '--masterExtractDir', dest='masterExtractDir', metavar='masterExtractDirectory', default=None, help='The name of the master directory sub-directory that contains the extract master CSV file and configuration specific to the extract')
    parser.add_argument ('-E', '--Extensive', dest='Extensive', action='store_true', help='Invoke Extensive checking to look for possible duplicates')
    parser.add_argument ('-m', '--masterDebugKey', dest='masterDebugKey', metavar='masterDebugKey', default=None, help='The key for triggering logging of information about a specific master record')
    parser.add_argument ('-n', '--masterDebugCount', dest='masterDebugCount', metavar='masterDebugCount', type=int, default=50000, help='A counter to trigger progress logging; a message every masterDebugCount(th) master record')
    parser.add_argument ('-q', '--quick', dest='quick', action='store_true', help='Quick check only of the master CSV file')
    parser.add_argument ('-v', '--verbose', dest='verbose', type=int, choices=range(0,5), help='The level of logging\n\t0=CRITICAL,1=ERROR,2=WARNING,3=INFO,4=DEBUG')
    parser.add_argument ('-o', '--logfile', dest='logfile', metavar='logfile', default=None, help='The name of a logging file')
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
    d.masterExtractDir = args.masterExtractDir
    d.Extensive = args.Extensive
    d.masterDebugKey = args.masterDebugKey
    d.masterDebugCount = args.masterDebugCount
    d.quick = args.quick

    # Read in the master configuration file
    f.getMasterConfig(False)

    # Read in the extract configuration file if required
    if d.masterExtractDir:
        # Read in the extract configuration file
        f.getMasterConfig(True)

    # Assemble the reporting columns
    d.reportingColumns = ['Date of Birth']
    d.reportingDates = ['Date of Birth']

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



    # Open Error file
    f.openErrorFile()

    # Open the raw master PMI file
    d.mc.masterOpenRawPMI()
    d.masterRecNo = 1
    ocount = 0

    # Create files for reporting family names and given names that should be checked
    d.date_style = NamedStyle(name='Date', number_format='dd-mmm-yyyy')
    f.openNameCheck('FamilyName')
    f.openNameCheck('GivenName')

    # Open the master.csv output file if we are saving it
    if not d.quick:
        try:
            if d.masterExtractDir:
                d.mfc = open(f'./{d.masterDir}/{d.masterExtractDir}/master.csv', 'wt', newline='')
            else:
                d.mfc = open(f'./{d.masterDir}/master.csv', 'wt', newline='')
        except:
            if d.masterExtractDir:
                logging.fatal('cannot create ./%s/%s/master.csv', d.masterDir, d.masterExtractDir)
            else:
                logging.fatal('cannot create ./%s/master.csv', d.masterDir)
            sys.exit(EX_CANTCREAT)
        d.mfcCSV = csv.writer(d.mfc, dialect='excel')
        d.mfcCSV.writerow(d.masterSaveTitles)

    # Process each raw PMI record (after cleaning up things link unprintables etc.)
    familyNamesChecked = 0
    familyNameErrors = 0
    givenNamesChecked = 0
    givenNameErrors = 0
    while d.mc.masterReadRawPMI():
        if not d.quick:
            d.mfcCSV.writerow(d.csvfields)
            ocount += 1

        # Save the alias and merge links for this record and check for unique PID and unique UR
        f.masterSaveLinks()

        thisUR = d.mc.masterCleanUR()        # Get the UR number

        # Check how we clean up the master family name and given name
        Mf = f.masterField('FamilyName')
        Mf = Mf.upper()
        Mf = re.sub('[A-Z \'-]', '', Mf)
        if Mf != '':
            familyNamesChecked += 1
            rep = []
            if d.ml.masterIsAlias():
                rep.append('[an alias]')
            elif d.ml.masterIsMerged() and (d.masterLinks['mergedIs'] == 'IN'):
                rep.append('[is merged]')
            else:
                rep.append('')
            rep.append(d.masterRecNo)
            rep.append(f.masterField('UR'))
            rep.append(f.masterField('FamilyName'))
            rep.append(d.mc.masterCleanFamilyName())
            Mf = d.mc.masterNeatFamilyName()
            Mf = Mf.upper()
            Mf = re.sub('[A-Z \'-]', '', Mf)
            if Mf != '':
                familyNameErrors += 1
                rep.append('Please check')
            d.worksheet['fnc'].append(rep)

        Mg = f.masterField('GivenName')
        Mg = Mg.upper()
        Mg = re.sub('[A-Z \'-]', '', Mg)
        if Mg != '':
            givenNamesChecked += 1
            rep = []
            if d.ml.masterIsAlias():
                rep.append('[an alias]')
            elif d.ml.masterIsMerged() and (d.masterLinks['mergedIs'] == 'IN'):
                rep.append('[is merged]')
            else:
                rep.append('')
            rep.append(d.masterRecNo)
            rep.append(f.masterField('UR'))
            rep.append(f.masterField('GivenName'))
            rep.append(d.mc.masterCleanGivenName())
            Mg = d.mc.masterNeatGivenName()
            Mg = Mg.upper()
            Mg = re.sub('[A-Z \'-]', '', Mg)
            if Mg != '':
                givenNameErrors += 1
                rep.append('Please check')
            d.worksheet['gnc'].append(rep)

        # If quick then this is all the checking we do
        if d.quick:
            # Report progress
            d.masterRecNo += 1
            if (d.masterRecNo % d.masterDebugCount) == 0:
                logging.info('%d records read', d.masterRecNo)
            continue

        # For the first pass we skip aliases - we're looking for duplicate
        # UR numbers associated with primary names.
        if d.ml.masterIsAlias():
            f.masterSetAlias()
            # Report progress
            d.masterRecNo += 1
            if (d.masterRecNo % d.masterDebugCount) == 0:
                logging.info('%d records read', d.masterRecNo)
            continue

        # And when checking for duplicates we skip merged patients
        if d.ml.masterIsMerged():
            f.masterSetMerged()
            if d.masterLinks['mergedIs'] == 'OUT':
                # Report progress
                d.masterRecNo += 1
                if (d.masterRecNo % d.masterDebugCount) == 0:
                    logging.info('%d records read', d.masterRecNo)
                continue

        # Next clean up the master family name and given name
        Mf = d.mc.masterCleanFamilyName()
        Mg = d.mc.masterCleanGivenName()
        Mdob = d.mc.masterCleanDOB()
        Msex = d.mc.masterCleanSex()
        thisKey = Mf + '~' + Mg + '~' + Mdob + '~' + Msex
        # This is the full key.
        # We use the UR number as the patient ID
        # i.e. fullKey[thisKey[]] is the patient UR number
        if thisKey not in d.fullKey:
            d.fullKey[thisKey] = []
        d.fullKey[thisKey].append(thisUR)

        # Collect Extensive checking data if required
        if d.Extensive:
            d.ExtensiveMasterRecKey[d.masterRecNo] = thisUR + '~' + thisKey + '~' + f.Sounds(Mf, Mg)
            if Mdob == '':
                d.ExtensiveMasterBirthdate[d.masterRecNo] = d.futureBirthdate
            else:
                d.ExtensiveMasterBirthdate[d.masterRecNo] = f.Birthdate(Mdob)
            if d.useMiddleNames:
                d.ExtensiveMasterMiddleNames[d.masterRecNo] = f.masterField('MiddleNames').upper()
            thisHash = {}
            for field in (sorted(d.ExtensiveFields.keys())):
                if f.masterField(field) == '':
                    thisHash[field] = None
                else:
                    thisHash[field] = hash(f.masterField(field))
            d.ExtensiveOtherMasterFields[d.masterRecNo] = thisHash

        # Report progress
        if (d.masterRecNo % d.masterDebugCount) == 0:
            logging.info('%d records read', d.masterRecNo)
        d.masterRecNo += 1

    # Close the raw master PMI file
    d.mc.masterCloseRawPMI()

    # If not quick then close the newly created secondary.csv file so that we can re-read it. If quick, then just exit
    if not d.quick:
        d.mfc.close()
    else:
        sys.exit(EX_OK)


    # Close files for reporting family names and given names that should be checked
    if d.masterExtractDir:
        fileName = f'./{d.masterDir}/{d.masterExtractDir}/{d.masterShortName}_FamilyNameCheck.xlsx'
    else:
        fileName = f'./{d.masterDir}/{d.masterShortName}_FamilyNameCheck.xlsx'
    f.PrintClose('fnc', 0, -1, fileName)

    if d.masterExtractDir:
        fileName = f'./{d.masterDir}/{d.masterExtractDir}/{d.masterShortName}_GivenNameCheck.xlsx'
    else:
        fileName = f'./{d.masterDir}/{d.masterShortName}_GivenNameCheck.xlsx'
    f.PrintClose('gnc', 0, -1, fileName)

    # Find the links for aliases and merged patients
    f.masterFindAliases()
    f.masterFindMerged()

    # Report any alias or merge link errors
    # We don't worry if the merged direction is 'IN' and the record merged in is missing as that can't cause a matching or find error
    if d.masterExtractDir:
        masterCSV = f'./{d.masterDir}/{d.masterExtractDir}/master.csv'
    else:
        masterCSV = f'./{d.masterDir}/master.csv'
    with open(masterCSV, 'rt') as csvfile:
        masterPMI = csv.reader(csvfile, dialect='excel')
        d.masterRecNo = 1
        heading = True
        for d.csvfields in masterPMI:
            if heading:
                heading = False
                continue
            thisUR = d.mc.masterCleanUR()
            if d.ml.masterIsAlias():
                if d.masterPrimRec[d.masterRecNo] is None:
                    alias = f.masterField('Alias')
                    if d.masterLinks['aliasLink'] == 'PID':
                        d.feCSV.writerow(['Alias with no matching primary id record', f'record No {d.masterRecNo}', f'{d.masterURname} {thisUR}', f'Alias of {d.masterPIDname} {alias}'])
                    else:
                        d.feCSV.writerow(['Alias with no matching primary id record', f'record No {d.masterRecNo}', f'{d.masterURname} {thisUR}', f'Alias of {d.masterURname} {alias}'])
            if d.ml.masterIsMerged() and (d.masterLinks['mergedIs'] != 'IN'):
                if d.masterNewRec[d.masterRecNo] is None:
                    if d.masterLinks['mergedLink'] == 'PID':
                        pid = f.masterField('Merged')
                        d.feCSV.writerow(['Merged to patient with no matching primary id record', f'record No {d.masterRecNo}', f'{d.masterURname} {thisUR}', f'Merged to {d.masterPIDname} {pid}'])
                    else:
                        ur = f.masterField('Merged')
                        d.feCSV.writerow(['Merged to patient with no matching primary id record', f'record No {d.masterRecNo}', f'{d.masterURname} {thisUR}', f'Merged to {d.masterURname} {ur}'])
            if (d.masterRecNo % d.masterDebugCount) == 0:
                logging.info('%d records processed', d.masterRecNo)
            d.masterRecNo += 1

    # Now look for probable duplicates
    probDuplicateChecks = 0
    f.openProbableDuplicatesCheck()
    for thisKey, thisFullKey in d.fullKey.items():
        if len(thisFullKey) > 1:
            (Mf, Mg, Mdob, Msex) = re.split('~', thisKey)
            probDuplicateChecks += 1
            # Output the probable duplicate UR numbers 20 at a time
            line = []
            for i, dupKey in enumerate(thisFullKey):
                if (i % 20) == 0:
                    if i == 0:
                        if Mdob == '':
                            line = ['Probable duplicate patients', Mf, Mg, Mdob, Msex]
                        else:
                            line = ['Probable duplicate patients', Mf, Mg, datetime.datetime.strptime(Mdob, '%Y-%m-%d'), Msex]
                    else:
                        d.worksheet['pdc'].append(line)
                        line = ['', '', '', '', '','']
                line.append(dupKey)
            d.worksheet['pdc'].append(line)

    if d.masterExtractDir:
        fileName = f'./{d.masterDir}/{d.masterExtractDir}/{d.masterShortName}_ProbableDuplicates.xlsx'
    else:
        fileName = f'./{d.masterDir}/{d.masterShortName}_ProbableDuplicates.xlsx'
    f.PrintClose('pdc', 0, 3, fileName)

    # Now look for possible duplicates - if the Extensive options is invoked
    possDuplicateChecks = 0
    if d.Extensive:
        f.openPossibleDuplicatesCheck()
        for rec1 in sorted(d.ExtensiveMasterRecKey):
            if d.ExtensiveMasterRecKey[rec1] == '':
                continue
            (UR1, Mf1, Mg1, Mdob1, Msex1, Mf1ny, Mf1dm,  Mf1sx, Mg1ny, Mg1dm, Mg1sx) = d.ExtensiveMasterRecKey[rec1].split('~')
            Birthdate1 = d.ExtensiveMasterBirthdate[rec1]
            OtherFields1 = d.ExtensiveOtherMasterFields[rec1]
            if d.useMiddleNames:
                Mm1 = d.ExtensiveMasterMiddleNames[rec1]
            possDuplicates = {}
            for rec2 in sorted(d.ExtensiveMasterRecKey):
                if rec2 <= rec1:
                    continue
                if d.ExtensiveMasterRecKey[rec2] == '':
                    continue
                (UR2, Mf2, Mg2, Mdob2, Msex2, Mf2ny, Mf2dm, Mf2sx, Mg2ny, Mg2dm, Mg2sx) = d.ExtensiveMasterRecKey[rec2].split('~')
                if (Mf1 == Mf2) and (Mg1 == Mg2) and (Mdob1 == Mdob2) and (Msex1 == Msex2):
                    continue
                Birthdate2 = d.ExtensiveMasterBirthdate[rec2]
                OtherFields2 = d.ExtensiveOtherMasterFields[rec2]
                if d.useMiddleNames:
                    Mm2 = d.ExtensiveMasterMiddleNames[rec2]
                weight = 1.0
                if Mf1 == Mf2:
                    soundFamilyNameConfidence = 100.0
                else:
                    (soundFamilyNameConfidence, weight) = f.FamilyNameSoundCheck(Mf1, Mf1ny, Mf1dm, Mf1sx, Mf2, Mf2ny, Mf2dm, Mf2sx, 1.0)
                if Mg1 == Mg2:
                    soundGivenNameConfidence = 100.0
                else:
                    (soundGivenNameConfidence, weight) = f.GivenNameSoundCheck(Mg1, Mg1ny, Mg1dm, Mg1sx, Mg2, Mg2ny, Mg2dm,  Mg2sx, 1.0)
                totalConfidence = 0
                totalWeight = 0
                for coreRoutine, (thisWeight, thisParam) in d.ExtensiveRoutines.items():
                    weight = 0.0
                    confidence = 0.0
                    if coreRoutine == 'FamilyName':
                        (confidence, weight) = f.FamilyNameCheck(Mf1, Mf2, thisWeight)
                    elif coreRoutine == 'FamilyNameSound':
                        (confidence, weight) = f.FamilyNameSoundCheck(Mf1, Mf1ny, Mf1dm, Mf1sx, Mf2, Mf2ny, Mf2dm, Mf2sx, thisWeight)
                    elif coreRoutine == 'GivenName':
                        (confidence, weight) = f.GivenNameCheck(Mg1, Mg2, thisWeight)
                    elif coreRoutine == 'GivenNameSound':
                        (confidence, weight) = f.GivenNameSoundCheck(Mg1, Mg1ny, Mg1dm, Mg1sx, Mg2, Mg2ny, Mg2dm, Mg2sx, thisWeight)
                    elif coreRoutine == 'MiddleNames':
                        if d.useMiddleNames:
                            (confidence, weight) = f.MiddleNamesCheck(Mm1, Mm2, thisWeight)
                    elif coreRoutine == 'MiddleNamesInitial':
                        if d.useMiddleNames:
                            (confidence, weight) = f.MiddleNamesInitialCheck(Mm1, Mm2, thisWeight)
                    elif coreRoutine == 'Sex':
                        (confidence, weight) = f.SexCheck(Msex1, Msex2, thisWeight)
                    elif coreRoutine == 'Birthdate':
                        (confidence, weight) = f.BirthdateCheck(Mdob1, Mdob2, thisWeight)
                    elif coreRoutine == 'BirthdateNearYear':
                        (confidence, weight) = f.BirthdateNearYearCheck(Mdob1, Mdob2, thisParam, thisWeight)
                    elif coreRoutine == 'BirthdateNearMonth':
                        (confidence, weight) = f.BirthdateNearMonthCheck(Mdob1, Mdob2, thisParam, thisWeight)
                    elif coreRoutine == 'BirthdateNearDay':
                        (confidence, weight) = f.BirthdateNearDayCheck(Birthdate1, Birthdate2, thisParam, thisWeight)
                    elif coreRoutine == 'BirthdateYearSwap':
                        (confidence, weight) = f.BirthdateYearSwapCheck(Mdob1, Mdob2, thisWeight)
                    elif coreRoutine == 'BirthdateDayMonthSwap':
                        (confidence, weight) = f.BirthdateDayMonthSwapCheck(Mdob1, Mdob2, thisWeight)
                    if weight > 0:
                        totalConfidence += confidence * weight
                        totalWeight += weight
                for field in (sorted(d.ExtensiveFields.keys())):
                    weight = d.ExtensiveFields[field]
                    if weight > 0:
                        if (OtherFields1[field] is not None) and (OtherFields2[field] is not None):
                            if OtherFields1[field] == OtherFields2[field]:
                                totalConfidence += 100.0 * weight
                                totalWeight += weight
                            else:
                                totalConfidence += 0.0 * weight
                                totalWeight += weight
                if totalWeight == 0:
                    continue
                totalConfidence = totalConfidence / totalWeight
                if totalConfidence >= d.ExtensiveConfidence:
                    if totalConfidence not in possDuplicates:
                        possDuplicates[totalConfidence] = []
                    if Mdob2 == '':
                        possDuplicates[totalConfidence].append([rec2, soundFamilyNameConfidence, soundGivenNameConfidence, UR2, Mf2, Mg2, Mdob2, Msex2])
                    else:
                        possDuplicates[totalConfidence].append([rec2, soundFamilyNameConfidence, soundGivenNameConfidence, UR2, Mf2, Mg2, Birthdate2, Msex2])
            if len(possDuplicates.keys()) > 0:
                possDuplicateChecks += 1
                if Mdob1 == '':
                    line = ['Possible duplcate patient for', UR1, Mf1, Mg1, Mdob1, Msex1]
                else:
                    line = ['Possible duplcate patient for', UR1, Mf1, Mg1, Birthdate1, Msex1]
                d.worksheet['pdc'].append(line)
                for confidence in (reversed(sorted(possDuplicates))):
                    for dup in (range(len(possDuplicates[confidence]))):
                        line = ['', '', '', '', '', '', confidence]
                        for field in (range(1, len(possDuplicates[confidence][dup]))):
                            line.append(possDuplicates[confidence][dup][field])
                        rec2 = possDuplicates[confidence][dup][0]
                        OtherFields2 = d.ExtensiveOtherMasterFields[rec2]
                        for field in (sorted(d.ExtensiveFields.keys())):
                            if (OtherFields1[field] is not None) and (OtherFields2[field] is not None):
                                if OtherFields1[field] == OtherFields2[field]:
                                    line.append('match')
                                else:
                                    line.append('')
                            else:
                                line.append('')
                        d.worksheet['pdc'].append(line)
            if ((rec1 + 1) % d.masterDebugCount) == 0:
                logging.info('%d records extensively checked', rec1 + 1)

        if d.masterExtractDir:
            fileName = f'./{d.masterDir}/{d.masterExtractDir}/{d.masterShortName}_PossibleDuplicates.xlsx'
        else:
            fileName = f'./{d.masterDir}/{d.masterShortName}_PossibleDuplicates.xlsx'
        f.PrintClose('pdc', 1, 4, fileName)

    # And finally, create the report
    f.openReport()
    heading = 'Phase 0 Testing - master PMI checks'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')
    heading = f'{d.masterLongName} PMI Status Report'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')
    d.rpt.write(f'{d.masterRawRecNo}\tRecords Read In\n')
    d.rpt.write(f'\t[{d.aCount} Alias records, {d.mCount} Merged records, {d.dCount} Deleted records]\n')
    d.rpt.write(f'{ocount}\tRecords Output for further testing (Deleted and bad records are not output)\n')
    if d.URdup > 0:
        d.rpt.write(f'{d.URdup}\tRecords have a duplicate {d.masterURname} number\n')
    if d.PIDdup > 0:
        d.rpt.write(f'{d.PIDdup}\tRecords have a duplicate {d.masterPIDname} number\n')
    if familyNamesChecked > 0:
        rptLine = f'{familyNamesChecked}\tFamily names modified (for comparisions)'
        if familyNameErrors > 0:
            rptLine += f' [{familyNameErrors} non-alphabeic family names]'
        rptLine += f' (file {d.masterShortName}_FamilyNameCheck.xlsx)\n'
        d.rpt.write(rptLine)
    if givenNamesChecked > 0:
        rptLine = f'{givenNamesChecked}\tGiven names modified (for comparisions)'
        if givenNameErrors > 0:
            rptLine += f' [{givenNameErrors} non-alphabeic given names]'
        rptLine += f' (file {d.masterShortName}_GivenNameCheck.xlsx)\n'
        d.rpt.write(rptLine)
    if probDuplicateChecks > 0:
        d.rpt.write(f'{probDuplicateChecks}\tProbable duplicates [same name/dob/sex - different {d.masterURname}] (file {d.masterShortName}_ProbableDuplicates.xlsx)\n')
    if d.Extensive:
        if possDuplicateChecks > 0:
            d.rpt.write(f'{possDuplicateChecks}\tPatients with possible duplicates (file {d.masterShortName}_PossibleDuplicates.xlsx)\n')
    d.rpt.close()

    # Close the error log csv file and exit
    d.fe.close()
    sys.exit(EX_OK)
