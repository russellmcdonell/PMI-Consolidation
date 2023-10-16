
# pylint: disable=line-too-long

'''
A script to check the goodness of health of a secondary PMI file

SYNOPSIS
$ python checkSecondary.py secondaryDirectory [-f secondaryExtractDirectory|--secondaryExtractDir=secondaryExtractDirectory]
                                              [-E|--Extensive] [-s secondaryDebugKey|--secondaryDebugKey=secondaryDebugKey]
                                              [-t secondaryDebugCount|--secondaryDebugCount=secondaryDebugCount]
                                              [-q|--quick] [-v loggingLevel|--verbose=loggingLevel] [-o logfile|--logfile=logfile]


OPTIONS
secondaryDirectory
The directory containing the secondary configuration (secondary.cfg) plus the  cleanSecondary.py and linkSecondary.py routines. Default is 'Secondary'
This directory may contain subdirectories where specific extracts will be found.

-f secondaryExtractDir|--secondaryExtractDir=secondaryExtractDir
The optional extract sub-directory, of the secondary directory, containing the extract secondary CSV file and parameters.

-E|--Extensive
Invoke extensive checking for possible duplicates. Extensive checking can invoke various function on the core data elements (FamilyName, GivenName, Birthdate, Sex)
plus simple checks of equality for any other field in the master PMI extract file.

-s secondaryDebugKey|--secondaryDebugKey=secondaryDebugKey
The key for triggering logging of information about a specific secondary record. Default is None

-t secondaryDebugCount|--secondaryDebugCount=secondaryDebugCount
A counter to trigger progress logging; a progress message is created every secondaryDebugCount(th) secondary record. Default is 50000

-q|--quick
Just performa a basic check of the secondary CSV file. Do not check alias or meged links. Do not create the cleaned up secondary CSV file.

-v loggingLevel|--verbose=loggingLevel
Set the level of logging that you want.

-o logfile|--logfile=logfile
The name of a log file where you want all messages captured.


THE MAIN CODE
Start by parsing the command line arguements and setting up logging.

Then read in the secondary PMI extract file and save it as a CSV file and make sure that there are no errors in the secondary PMI extract.
Many of the checks are contained in secondaryDirectory/cleanSecondary.py
secondaryDirectory/cleanSecondary.py may need editing to match the constraints for valid secondary records
Checks include records with the wrong number of fields, the UR number is valid, alias records have a matching secondary record,
merged records have a matching secondary record and probable duplicates.
This is based up cleaned up family name and cleaned up given name so there is no guaranttees that they are duplicates.
No checks of address, medicare number, next of kin or any other secondary identifiers are attempted unless the Extensive options is involked.
If the Extensive options is involked then a further check is conducted for possible duplicates.

The function in secondaryDirectory/cleanSecondary.py associated with cleaning up family names, given names, birthdates and gender
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
Then read in the secondary file and save it as a CSV file, making sure that there are no errors in the secondary PMI extract.
    '''

    # Save the program name
    d.progName = sys.argv[0]
    d.progName = d.progName[0:-3]        # Strip off the .py ending
    d.scriptType = 'secondary'

    # Get the options
    parser = argparse.ArgumentParser(description='Check the goodness of health of a secondary PMI extraction')
    parser.add_argument ('secondaryDir', metavar='secondaryDirectory', help='The name of directory containg the secondary configuration and cleanSecondary.py routines')
    parser.add_argument ('-f', '--secondaryExtractDir', dest='secondaryExtractDir', metavar='secondaryExtractDirectory', default=None, help='The name of the secondary directory sub-directory that contains the extract secondary CSV file and configuration specific to the extract')
    parser.add_argument ('-E', '--Extensive', dest='Extensive', action='store_true', help='Invoke Extensive checking to look for possible duplicates')
    parser.add_argument ('-s', '--secondaryDebugKey', dest='secondaryDebugKey', metavar='secondaryDebugKey', default=None, help='The key for triggering logging of information about a specific secondary record')
    parser.add_argument ('-t', '--secondaryDebugCount', dest='secondaryDebugCount', metavar='secondaryDebugCount', type=int, default=50000, help='A counter to trigger progress logging; a message every secondaryDebugCount(th) secondary record')
    parser.add_argument ('-q', '--quick', dest='quick', action='store_true', help='Just a basic check of the secondary CSV file')
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


    d.secondaryDir = args.secondaryDir
    d.secondaryExtractDir = args.secondaryExtractDir
    d.Extensive = args.Extensive
    d.secondaryDebugKey = args.secondaryDebugKey
    d.secondaryDebugCount = args.secondaryDebugCount
    d.quick = args.quick

    # Read in the secondary configuration file
    f.getSecondaryConfig(False)

    # Read in the extract configuration file if required
    if d.secondaryExtractDir:
        # Read in the extract configuration file
        f.getSecondaryConfig(True)

    # Assemble the reporting columns
    d.reportingColumns = ['Date of Birth']
    d.reportingDates = ['Date of Birth']

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

    # Read in the Not Matched PIDs
    f.getNotMatched()

    # Open the raw secondary PMI file
    d.sc.secondaryOpenRawPMI()
    d.secondaryRecNo = 1
    ocount = 0

    # Create files for reporting family names and given names that should be checked
    d.date_style = NamedStyle(name='Date', number_format='dd-mmm-yyyy')
    f.openNameCheck('FamilyName')
    f.openNameCheck('GivenName')

    # Open the secondary.csv output file if we are saving it
    if not d.quick:
        try:
            if d.secondaryExtractDir:
                d.sfc = open(f'./{d.secondaryDir}/{d.secondaryExtractDir}/secondary.csv', 'wt', newline='')
            else:
                d.sfc = open(f'./{d.secondaryDir}/secondary.csv', 'wt', newline='')
        except:
            if d.secondaryExtractDir:
                logging.fatal('cannot create ./%s/%s/secondary.csv', d.secondaryDir, d.secondaryExtractDir)
            else:
                logging.fatal('cannot create ./%s/secondary.csv', d.secondaryDir)
            sys.exit(EX_CANTCREAT)
        d.sfcCSV = csv.writer(d.sfc, dialect='excel')
        d.sfcCSV.writerow(d.secondarySaveTitles)

    # Process each raw PMI record (after cleaning up things link unprintables etc.)
    familyNamesChecked = 0
    familyNameErrors = 0
    givenNamesChecked = 0
    givenNameErrors = 0
    while d.sc.secondaryReadRawPMI():
        if not d.quick:
            d.sfcCSV.writerow(d.csvfields)
            ocount += 1

        # Save the alias and merge links for this record and check for unique PID and unique UR
        f.secondarySaveLinks()

        thisUR = d.sc.secondaryCleanUR()        # Get the UR number

        # Check how we clean up the secondary family name and given name
        Mf = f.secondaryField('FamilyName')
        Mf = Mf.upper()
        Mf = re.sub('[A-Z \'-]', '', Mf)
        if Mf != '':
            familyNamesChecked += 1
            rep = []
            if d.sl.secondaryIsAlias():
                rep.append('[an alias]')
            elif d.sl.secondaryIsMerged() and (d.secondaryLinks['mergedIs'] == 'IN'):
                rep.append('[is merged]')
            else:
                rep.append('')
            rep.append(d.secondaryRecNo)
            rep.append(f.secondaryField('UR'))
            rep.append(f.secondaryField('FamilyName'))
            rep.append(d.sc.secondaryCleanFamilyName())
            Mf = d.sc.secondaryNeatFamilyName()
            Mf = Mf.upper()
            Mf = re.sub('[A-Z \'-]', '', Mf)
            if Mf != '':
                familyNameErrors += 1
                rep.append('Please check')
            d.worksheet['fnc'].append(rep)

        Mg = f.secondaryField('GivenName')
        Mg = Mg.upper()
        Mg = re.sub('[A-Z \'-]', '', Mg)
        if Mg != '':
            givenNamesChecked += 1
            rep = []
            if d.sl.secondaryIsAlias():
                rep.append('[an alias]')
            elif d.sl.secondaryIsMerged() and (d.secondaryLinks['mergedIs'] == 'IN'):
                rep.append('[is merged]')
            else:
                rep.append('')
            rep.append(d.secondaryRecNo)
            rep.append(f.secondaryField('UR'))
            rep.append(f.secondaryField('GivenName'))
            rep.append(d.sc.secondaryCleanGivenName())
            Mg = d.sc.secondaryNeatGivenName()
            Mg = Mg.upper()
            Mg = re.sub('[A-Z \'-]', '', Mg)
            if Mg != '':
                givenNameErrors += 1
                rep.append('Please check')
            d.worksheet['gnc'].append(rep)

        # If quick then this is all the checking we do
        if d.quick:
            # Report progress
            d.secondaryRecNo += 1
            if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                logging.info('%d records read', d.secondaryRecNo)
            continue

        # For the first pass we skip aliases - we're looking for duplicate
        # UR numbers associated with primary names.
        if d.sl.secondaryIsAlias():
            f.secondarySetAlias()
            # Report progress
            d.secondaryRecNo += 1
            if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                logging.info('%d records read', d.secondaryRecNo)
            continue

        # And for checking for probable duplicates we skip merged patients
        if d.sl.secondaryIsMerged():
            f.secondarySetMerged()
            if d.secondaryLinks['mergedIs'] == 'OUT':
                # Report progress
                d.secondaryRecNo += 1
                if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                    logging.info('%d records read', d.secondaryRecNo)
                continue

        # Next clean up the secondary family name and given name
        Sf = d.sc.secondaryCleanFamilyName()
        Sg = d.sc.secondaryCleanGivenName()
        Sdob = d.sc.secondaryCleanDOB()
        Ssex = d.sc.secondaryCleanSex()
        thisKey = Sf + '~' + Sg + '~' + Sdob + '~' + Ssex
        # This is the full key.
        # We use the UR number as the patient ID
        # i.e. fullKey[thisKey[]] is the patient UR number
        if thisKey not in d.fullKey:
            d.fullKey[thisKey] = []
        d.fullKey[thisKey].append(thisUR)

        # Collect Extensive checking data if required
        if d.Extensive:
            d.ExtensiveSecondaryRecKey[d.secondaryRecNo] = thisUR + '~' + thisKey + '~' + f.Sounds(Sf, Sg)
            if Sdob == '':
                d.ExtensiveSecondaryBirthdate[d.secondaryRecNo] = d.futureBirthdate
            else:
                d.ExtensiveSecondaryBirthdate[d.secondaryRecNo] = f.Birthdate(Sdob)
            if d.useMiddleNames:
                d.ExtensiveSecondaryMiddleNames[d.secondaryRecNo] = f.secondaryField('MiddleNames').upper()
            thisHash = {}
            for field in (sorted(d.ExtensiveFields.keys())):
                if f.secondaryField(field) == '':
                    thisHash[field] = None
                else:
                    thisHash[field] = hash(f.secondaryField(field))
            d.ExtensiveOtherSecondaryFields[d.secondaryRecNo] = thisHash

        # Report progress
        if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
            logging.info('%d records read', d.secondaryRecNo)
        d.secondaryRecNo += 1

    # Close the raw secondary PMI file
    d.sc.secondaryCloseRawPMI()

    # If not quick then close the newly created secondary.csv file so that we can re-read it. If quick, then just exit
    if not d.quick:
        d.sfc.close()
    else:
        sys.exit(EX_OK)


    # Close files for reporting family names and given names that should be checked
    if d.secondaryExtractDir:
        fileName = f'./{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_FamilyNameCheck.xlsx'
    else:
        fileName = f'./{d.secondaryDir}/{d.secondaryShortName}_FamilyNameCheck.xlsx'
    f.PrintClose('fnc', 0, -1, fileName)

    if d.secondaryExtractDir:
        fileName = f'./{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_GivenNameCheck.xlsx'
    else:
        fileName = f'./{d.secondaryDir}/{d.secondaryShortName}_GivenNameCheck.xlsx'
    f.PrintClose('gnc', 0, -1, fileName)

    # Find the links for aliases and merged patients
    f.secondaryFindAliases()
    f.secondaryFindMerged()

    # Report any alias or merge link errors
    # We don't worry if the merged direction is 'IN' and the record merged in is missing as that can't cause a matching or find error
    if d.secondaryExtractDir:
        secondaryCSV = f'./{d.secondaryDir}/{d.secondaryExtractDir}/secondary.csv'
    else:
        secondaryCSV = f'./{d.secondaryDir}/secondary.csv'
    with open(secondaryCSV, 'rt') as csvfile:
        secondaryPMI = csv.reader(csvfile, dialect='excel')
        d.secondaryRecNo = 1
        heading = True
        for d.csvfields in secondaryPMI:
            if heading:
                heading = False
                continue
            thisUR = d.sc.secondaryCleanUR()
            if d.sl.secondaryIsAlias():
                if d.secondaryPrimRec[d.secondaryRecNo] is None:
                    alias = f.secondaryField('Alias')
                    if d.secondaryLinks['aliasLink'] == 'PID':
                        d.feCSV.writerow(['Alias with no matching primary id record', 'record No {d.secondaryRecNo}', f'{d.secondaryURname} {thisUR}', f'Alias of {d.secondaryPIDname} {alias}'])
                    else:
                        d.feCSV.writerow(['Alias with no matching primary id record', 'record No {d.secondaryRecNo}', f'{d.secondaryURname} {thisUR}', f'Alias of {d.secondaryURname} {alias}'])
            if d.sl.secondaryIsMerged() and (d.secondaryLinks['mergedIs'] != 'IN'):
                if d.secondaryNewRec[d.secondaryRecNo] is None:
                    if d.secondaryLinks['mergedLink'] == 'PID':
                        pid = f.secondaryField('Merged')
                        d.feCSV.writerow(['Merged to patient with no matching primary id record', f'record No {d.secondaryRecNo}', f'{d.secondaryURname} {thisUR}', f'Merged to {d.secondaryPIDname} {pid}'])
                    else:
                        ur = f.secondaryField('Merged')
                        d.feCSV.writerow(['Merged to patient with no matching primary id record', f'record No {d.secondaryRecNo}', f'{d.secondaryURname} {thisUR}', f'Merged to {d.secondaryURname} {ur}'])
            if (d.secondaryRecNo % d.secondaryDebugCount) == 0:
                logging.info('%d records processed', d.secondaryRecNo)
            d.secondaryRecNo += 1

    # Now look for probable duplicates
    probDuplicateChecks = 0
    f.openProbableDuplicatesCheck()
    for thisKey, thisFullKey in d.fullKey.items():
        if len(thisFullKey) > 1:
            probDuplicateChecks += 1
            # Output the probable duplicate UR numbers 20 at a time
            line = []
            for i, eachDup in enumerate(d.fullKey[thisKey]):
                if (i % 20) == 0:
                    if i == 0:
                        (Sf, Sg, Sdob, Ssex) = re.split('~', thisKey)
                        if Sdob == '':
                            line = ['Probable duplicate patients', Sf, Sg, Sdob, Ssex]
                        else:
                            line = ['Probable duplicate patients', Sf, Sg, datetime.datetime.strptime(Sdob, '%Y-%m-%d'), Ssex]
                    else:
                        d.worksheet['pdc'].append(line)
                        line = ['', '', '', '', '','']
                line.append(eachDup)
            d.worksheet['pdc'].append(line)

    if d.secondaryExtractDir:
        fileName = f'./{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_ProbableDuplicates.xlsx'
    else:
        fileName = f'./{d.secondaryDir}/{d.secondaryShortName}_ProbableDuplicates.xlsx'
    f.PrintClose('pdc', 0, 3, fileName)

    # Now look for possible duplicates - if the Extensive options is invoked
    possDuplicateChecks = 0
    if d.Extensive:
        f.openPossibleDuplicatesCheck()
        for rec1 in sorted(d.ExtensiveSecondaryRecKey):
            if d.ExtensiveSecondaryRecKey[rec1] == '':
                continue
            (UR1, Sf1, Sg1, Sdob1, Ssex1, Sf1ny, Sf1dm, Sf1sx, Sg1ny, Sg1dm, Sg1sx) = d.ExtensiveSecondaryRecKey[rec1].split('~')
            Birthdate1 = d.ExtensiveSecondaryBirthdate[rec1]
            OtherFields1 = d.ExtensiveOtherSecondaryFields[rec1]
            if d.useMiddleNames:
                Sm1 = d.ExtensiveSecondaryMiddleNames[rec1]
            possDuplicates = {}
            for rec2 in sorted(d.ExtensiveSecondaryRecKey):
                if rec2 <= rec1:
                    continue
                if d.ExtensiveSecondaryRecKey[rec2] == '':
                    continue
                (UR2, Sf2, Sg2, Sdob2, Ssex2, Sf2ny, Sf2dm, Sf2sx, Sg2ny, Sg2dm, Sg2sx) = d.ExtensiveSecondaryRecKey[rec2].split('~')
                if (Sf1 == Sf2) and (Sg1 == Sg2) and (Sdob1 == Sdob2) and (Ssex1 == Ssex2):
                    continue
                Birthdate2 = d.ExtensiveSecondaryBirthdate[rec2]
                OtherFields2 = d.ExtensiveOtherSecondaryFields[rec2]
                if d.useMiddleNames:
                    Sm2 = d.ExtensiveSecondaryMiddleNames[rec2]
                weight = 1.0
                if Sf1 == Sf2:
                    soundFamilyNameConfidence = 100.0
                else:
                    (soundFamilyNameConfidence, weight) = f.FamilyNameSoundCheck(Sf1, Sf1ny, Sf1dm, Sf1sx, Sf2, Sf2ny, Sf2dm, Sf2sx, 1.0)
                if Sg1 == Sg2:
                    soundGivenNameConfidence = 100.0
                else:
                    (soundGivenNameConfidence, weight) = f.GivenNameSoundCheck(Sg1, Sg1ny, Sg1dm, Sg1sx, Sg2, Sg2ny, Sg2dm, Sg2sx, 1.0)
                totalConfidence = 0
                totalWeight = 0
                for coreRoutine, (thisWeight, thisParam) in d.ExtensiveRoutines.items():
                    weight = 0.0
                    confidence = 0.0
                    if coreRoutine == 'FamilyName':
                        (confidence, weight) = f.FamilyNameCheck(Sf1, Sf2, thisWeight)
                    elif coreRoutine == 'FamilyNameSound':
                        (confidence, weight) = f.FamilyNameSoundCheck(Sf1, Sf1ny, Sf1dm, Sf1sx, Sf2, Sf2ny, Sf2dm, Sf2sx, thisWeight)
                    elif coreRoutine == 'GivenName':
                        (confidence, weight) = f.GivenNameCheck(Sg1, Sg2, thisWeight)
                    elif coreRoutine == 'GivenNameSound':
                        (confidence, weight) = f.GivenNameSoundCheck(Sg1, Sg1ny, Sg1dm, Sg1sx, Sg2, Sg2ny, Sg2dm, Sg2sx, thisWeight)
                    elif coreRoutine == 'MiddleNames':
                        if d.useMiddleNames:
                            (confidence, weight) = f.MiddleNamesCheck(Sm1, Sm2, thisWeight)
                    elif coreRoutine == 'MiddleNamesInitial':
                        if d.useMiddleNames:
                            (confidence, weight) = f.MiddleNamesInitialCheck(Sm1, Sm2, thisWeight)
                    elif coreRoutine == 'Sex':
                        (confidence, weight) = f.SexCheck(Ssex1, Ssex2, thisWeight)
                    elif coreRoutine == 'Birthdate':
                        (confidence, weight) = f.BirthdateCheck(Sdob1, Sdob2, thisWeight)
                    elif coreRoutine == 'BirthdateNearYear':
                        (confidence, weight) = f.BirthdateNearYearCheck(Sdob1, Sdob2, thisParam, thisWeight)
                    elif coreRoutine == 'BirthdateNearMonth':
                        (confidence, weight) = f.BirthdateNearMonthCheck(Sdob1, Sdob2, thisParam, thisWeight)
                    elif coreRoutine == 'BirthdateNearDay':
                        (confidence, weight) = f.BirthdateNearDayCheck(Birthdate1, Birthdate2, thisParam, thisWeight)
                    elif coreRoutine == 'BirthdateYearSwap':
                        (confidence, weight) = f.BirthdateYearSwapCheck(Sdob1, Sdob2, thisWeight)
                    elif coreRoutine == 'BirthdateDayMonthSwap':
                        (confidence, weight) = f.BirthdateDayMonthSwapCheck(Sdob1, Sdob2, thisWeight)
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
                    if Sdob2 == '':
                        possDuplicates[totalConfidence].append([rec2, soundFamilyNameConfidence, soundGivenNameConfidence, UR2, Sf2, Sg2, Sdob2, Ssex2])
                    else:
                        possDuplicates[totalConfidence].append([rec2, soundFamilyNameConfidence, soundGivenNameConfidence, UR2, Sf2, Sg2, Birthdate2, Ssex2])
            if len(possDuplicates.keys()) > 0:
                possDuplicateChecks += 1
                if Sdob1 == '':
                    line = ['Possible duplcate patient for', UR1, Sf1, Sg1, Sdob1, Ssex1]
                else:
                    line = ['Possible duplcate patient for', UR1, Sf1, Sg1, Birthdate1, Ssex1]
                d.worksheet['pdc'].append(line)
                for confidence in (reversed(sorted(possDuplicates))):
                    for dup in (range(len(possDuplicates[confidence]))):
                        line = ['', '', '', '', '', '', confidence]
                        for field in (range(1, len(possDuplicates[confidence][dup]))):
                            line.append(possDuplicates[confidence][dup][field])
                        rec2 = possDuplicates[confidence][dup][0]
                        OtherFields2 = d.ExtensiveOtherSecondaryFields[rec2]
                        for field in (sorted(d.ExtensiveFields.keys())):
                            if (OtherFields1[field] is not None) and (OtherFields2[field] is not None):
                                if OtherFields1[field] == OtherFields2[field]:
                                    line.append('match')
                                else:
                                    line.append('')
                            else:
                                line.append('')
                        d.worksheet['pdc'].append(line)
            if ((rec1  + 1) % d.secondaryDebugCount) == 0:
                logging.info('%d records extensively checked', rec1 + 1)

        if d.secondaryExtractDir:
            fileName = f'./{d.secondaryDir}/{d.secondaryExtractDir}/{d.secondaryShortName}_PossibleDuplicates.xlsx'
        else:
            fileName = f'./{d.secondaryDir}/{d.secondaryShortName}_PossibleDuplicates.xlsx'
        f.PrintClose('pdc', 1, 4, fileName)


    # And finally, create the report
    f.openReport()
    heading = 'Phase 0 Testing - secondary PMI checks'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')
    heading = f'{d.secondaryLongName} PMI Status Report'
    d.rpt.write(f'{heading}\n')
    heading = re.sub('[^ ]', '_', heading)
    d.rpt.write(f'{heading}\n')
    d.rpt.write(f'{d.secondaryRawRecNo}\tRecords Read In\n')
    d.rpt.write(f'\t[{d.aCount} Alias records, {d.mCount} Merged records, {d.dCount} Deleted records]\n')
    d.rpt.write(f'{ocount}\tRecords Output for further testing (Deleted and bad records are not output)\n')
    if d.URdup > 0:
        d.rpt.write(f'{d.URdup}\tRecords have a duplicate {d.secondaryURname} number\n' )
    if d.PIDdup > 0:
        d.rpt.write(f'{d.PIDdup}\tRecords have a duplicate {d.secondaryPIDname} number\n')
    if d.noAltURcount > 0:
        d.rpt.write(f'{d.noAltURcount}\tRecords do not have a {d.secondaryAltURname} number\n')
    if d.dudAltURcount > 0:
        d.rpt.write(f'{d.dudAltURcount}\tRecords have a invalid {d.secondaryAltURname} number\n')
    if d.AltURdup > 0:
        d.rpt.write(f'{d.AltURdup}\tRecords have a duplicate {d.secondaryAltURname} number\n')
    if familyNamesChecked > 0:
        rptLine = f'{familyNamesChecked}\tFamily name modified (for comparisions)'
        if familyNameErrors > 0:
            rptLine += f' [{familyNameErrors} non-alphabeic family names]'
        rptLine += f' (file {d.secondaryShortName}_FamilyNameCheck.xlsx)\n'
        d.rpt.write(rptLine)
    if givenNamesChecked > 0:
        rptLine = f'{givenNamesChecked}\tGiven names modified (for comparisions)'
        if givenNameErrors > 0:
            rptLine += f' [{givenNameErrors} non-alphabeic given names]'
        rptLine += f' (file {d.secondaryShortName}_GivenNameCheck.xlsx)\n'
        d.rpt.write(rptLine)
    if probDuplicateChecks > 0:
        d.rpt.write(f'{probDuplicateChecks}\tProbable duplicates [same name/dob/sex - different {d.secondaryURname}] (file {d.secondaryShortName}_ProbableDuplicates.xlsx)\n')
    if possDuplicateChecks > 0:
        d.rpt.write(f'{possDuplicateChecks}\tPossible duplicates (file {d.secondaryShortName}_PossibleDuplicates.xlsx)\n')
    d.rpt.close()

    # Close the error log csv file and exit
    d.fe.close()
    sys.exit(EX_OK)
    