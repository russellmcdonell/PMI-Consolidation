"""The common data"""

# pylint: disable=invalid-name, line-too-long

import datetime

progName = None            # The name of the script - must be set by each script

# The command line arguements
masterDir = ''            # Master directory
masterExtractDir = ''        # Master extract directory
secondaryDir = ''        # Secondary directory
secondaryExtractDir = ''    # Secondary extract directory
Extensive=False            # Invoke Extensive checking
masterDebugKey = ''        # Master Debug Key
masterDebugCount = 50000    # Master Debug Count
secondaryDebugKey = ''        # Secondary Debug Key
secondaryDebugCount = 10000    # Secondary Debug Count
quick = False            # Quick - report file health only

scriptType = ''            # The name of the script

# From Section [masterFile] in either master.cfg or extract.cfg
masterFileName = None        # Name of extract file
masterFieldCount = None        # Number of column in the extract file

# From Section [masterSaveColumns] in either master.cfg or extract.cfg
masterSaveColumns = []    # Column number (counting from 0) of the columns in the extract file that will be saved in master.csv
masterSaveTitles = []        # Column names/titles for the columns (from the extract file) in master.csv

# From Section [masterNames] in master.cfg
masterShortName = None        # master PMI short name
masterLongName = None        # master PMI long name

# From Section [masterIDs] in master.cfg
masterPIDname = None        # master PMI PID name
masterURname = None        # master PMI UR name

# From Section [masterHas] in master.cfg
masterHas = {}            # Cleaned up master PMI - Keys: concepts, Values: associated column names
masterIs = {}            # Cleaned up master PMI - Keys: column names, Values: column numbers

# From Section [masterLinks] in master.cfg
masterLinks = {}        # Linkage information for merges, aliases and deleted PMI records

# From Section [masterReporting] in master.cfg
masterReportingColumns = []    # Cleaned up master PMI reporting concepts

mc = None            # The name space for the masterDirectory/cleanMaster.py subroutines
ml = None            # The name space for the masterDir/linkMaster.py subroutines

mfr = None            # File handle for reading a raw master PMI file
dialect = None            # The csv dialect of a csv file
masterHasHeader = False        # Flag to indicate that a raw master PMI file has a header record
masterIsCSV = True        # Flag to indicate that a raw master PMI file is a CSV file
mwb = None              # Master Workbook (for Excel extracts)
mws = None              # Master Worksheet (for Excel exracts)
mws_iter_rows = None    # Master rows generator (for Excel extracts)
mfc = None            # File handle for reading/writing the cleaned up master PMI file
masterRawRecNo = 0        # Record number of raw record read in from to Master PMI extract file
masterRecNo = 0            # Record number of record read in from/written to cleaned up Master PMI file
URrec = {}            # Record Number for each UR - Keys: UR, Values: masterRecNo
PIDrec = {}            # Record Number for each PID - Keys: PID, Values: masterRecNo
URasIntRec = {}            # Record Number for the integer representation each UR - Keys: intUR, Values: masterRecNo
AltURrecNo = {}            # Record Number for each Alt UR in secondary PMI- Keys: AltUR, Values: secondaryRecNo

fe = None            # File handle for writing the error messages
feCSV = None            # CSV Writer for writing the error records
rpt = None            # File handle for writing the report

workbook = {}            # Workbooks for reporting
worksheet = {}            # Worksheets for reporting

# From Section [secondaryPMInames] in secondary.cfg
secondaryShortName = None    # secondary PMI short name
secondaryLongName = None    # secondary PMI long name
secondaryPIDname = None        # secondary PMI PID name

# From Section [secondaryIDs] in secondary.cfg
secondaryPIDname = None        # secondary PMI PID name
secondaryURname = None        # secondary PMI UR name
secondaryAltURname = None    # secondary PMI Alternate UR name

# From Section [secondaryHas] in secondary.cfg
secondaryHas = {}        # Cleaned up secondary PMI - Keys: concepts, Values: associated column names
secondaryIs = {}        # Cleaned up secondary PMI - Keys: column names, Values: column numbers

# From Section [secondaryLinks] in secondary.cfg
secondaryLinks = {}        # Linkage information for merges, aliases and deleted PMI records

# From Section [secondaryReporting] in secondary.cfg
secondaryReportingColumns = []    # Cleaned up secondary PMI reporting concepts

# The combined reporting columns (used in matchAltUR and findUR)
reportingColumns = []
reportingDates = []
date_style = None

# From Section [extract] in either secondary.cfg or extract.cfg
secondaryFileName = None    # Name of extract file
secondaryFieldCount = None    # Number of column in the extract file
secondarySaveColumns = []    # Column number (counting from 0) of the columns in the extract file that will be saved in secondary.csv
secondarySaveTitles = []    # Column names/titles for the columns (from the extract file) in secondary.csv

ExtensiveConfidence=100.0        # The confidence level that must be reached before possible duplicates/matches are reported
                    # The valid Extensive checking routines
DefinedRoutines = ['FamilyName', 'FamilyNameSound', 'GivenName', 'GivenNameSound', 'MiddleNames', 'MiddleNamesInitial', 'Sex', 'Birthdate', 'BirthdateNearYear', 'BirthdateNearMonth', 'BirthdateNearDay', 'BirthdateYearSwap', 'BirthdateDayMonthSwap']
ExtensiveRoutines = {}            # The configured Extensive checking routines - their weights and possibly their parameter
ExtensiveFields = {}            # The weights for the 'other fields'
useMiddleNames = False            # Middle names are being checked
ExtensiveMasterRecKey = {}        # The record number and key for each master record
ExtensiveMasterMiddleNames = {}        # the record number and Middle Names for each master record
ExtensiveMasterBirthdate = {}        # the record number and Birthdate for each master record
ExtensiveOtherMasterFields = {}        # The record number and hash of the other fields for each master record
ExtensiveSecondaryRecKey = {}        # The record number and key for each secondary record
ExtensiveSecondaryMiddleNames = {}    # the record number and Middle Names for each secondary record
ExtensiveSecondaryBirthdate = {}    # the record number and Birthdate for each secondary record
ExtensiveOtherSecondaryFields = {}    # The record number and hash of the other fields for each secondary record

sc = None            # The name space for the secondaryDir/Clean%secondaryShortName%.py subroutines
sl = None            # The name space for the secondaryDir/Link%secondaryShortName%.py subroutines

sfr = None            # File handle for reading a raw secondary PMI file
sfrCSV = None            # The csv reader object for reading the raw secondary PMI file
secondaryHasHeader = False    # Flag to indicate that a raw secondary PMI file has a header record
secondaryIsCSV = True        # Flag to indicate that a raw secondary PMI file is a CSV file
swb = None              # Secondary Workbook (for Excel extracts)
sws = None              # Secondary Worksheet (for Excel exracts)
sws_iter_rows = None    # Secondary rows generator (for Excel extracts)
sfc = None            # File handle for reading/writing the cleaned up secondary PMI file
sfcCSV = None            # The csv reader object for reading/writing the cleaned up secondary PMI file

today = datetime.date.today()    # Today's data
futureBirthdate = datetime.date(2100, 1, 1)    # Missing birthdate data value
dCount = 0            # Deleted patient records
aCount = 0            # Alias patient records
mCount = 0            # Merged patient records
URdup = 0            # Count of records with duplicate UR numbers
PIDdup = 0            # Count of records with duplicate PID values
AltURdup = 0            # Count of records with duplicate Alt UR numbers
noAltURcount = 0        # Count of secondary PMI records with no Alt UR
dudAltURcount = 0        # Count of secondary PMI records with a dud Alt UR
matchVolume = 1            # Count of matched files created
matchDupVolume = 1        # Count of duplicate records files created
mmatch = 0            # Count of secondary records matched to a merged master PMI patient
amatch = 0            # Count of secondary records matched to an alias master PMI patient
match = 0            # Count of secondary records matched to only one master PMI patient
fncmis = 0            # Count of secondary records matched, but family not matched - close
fncmisdn = 0            # Count of secondary records matched, but family not matched - close (Done)
fnomis = 0            # Count of secondary records matched, but family not matched - only
fnomisdn = 0            # Count of secondary records matched, but family not matched - only (Done)
fnpmis = 0            # Count of secondary records matched, but family not matched - plus
fnpmisdn = 0            # Count of secondary records matched, but family not matched - plus (Done)
dobmis = 0            # Count of secondary records matched, but birthdate did not match
dobmisdn = 0            # Count of secondary records matched, but birthdate did not match (Done)
gnmis = 0            # Count of secondary records matched, but given name did not match
gnmisdn = 0            # Count of secondary records matched, but given name did not match (Done)
sexmis = 0            # Count of secondary records matched, but sex did not match
sexmisdn = 0            # Count of secondary records matched, but sex did not match (Done)
extensivematch = 0        # Count of secondary records matched exactly using extensive matching
probablematch = 0        # Count of secondary records matched approximately using extensive matching

foundDoneVolume = 1        # Count of found_Done files created
foundToDoVolume = 1        # Count of found_ToDo files created
mfound = 0            # Count of secondary records found to be a merged master PMI patent
afound = 0            # Count of secondary records found to be an alias master PMI patent
URrecNos = []            # List of multiple found master records
URrecSounds = []        # List of matching sound for multipl found master records
foundtd = 0            # Count of secondary records found
founddn = 0            # Count of secondary records found (Done)
dfoundtd = 0            # Count of secondary records with duplicate master records found
dfounddn = 0            # Count of secondary records with duplicate master records found (Done)
dpfoundtd = 0            # Count of secondary records with duplicate similar master records found
dpfounddn = 0            # Count of secondary records with duplicate similar master records found (Done)
extensivefound = 0        # Count of identical secondary records found using extensive matching
probablefound = 0        # Count of similar secondary records found using extensive matching


csvfields = []            # an array for holding fields from/to CSV files
fullKey = {}            # The Full Key / secondary PMI the record number(s) for this key
secKey = []            # The Key for each secondary PMI record
secondaryPID = []        # The record number for this secondary PID
keySsx = {}            # The Sounds Like (Soundex) Keys and the record number for this key
keySdm = {}            # The Sounds Like (Double Metaphone) Keys and the record number for this key
keySny = {}            # The Sounds Like (NYSIIS) Keys and the record number for this key
key123 = {}            # The Family Name, Sex and DOB Key and the rec. number
key124 = {}            # The Family Name, Sex and Given Name Key and the rec. no.
key134 = {}            # The Family Name, DOB and Given Name Key and the rec.rd no.
key234 = {}            # The Sex, DOB and Given Name Key and the record number
masterNewRec = {}        # Record number of the "merged TO" patient
masterLinkRec = {}        # Record number of the "merged INTO" patient
masterPrimRec = {}        # Record number of primary for an alias
secondaryNewRec = {}        # Record number of the "merged TO" patient
secondaryLinkRec = {}        # Record number of the "merged INTO" patient
secondaryPrimRec = {}        # Record number of primary for an alias
altUR = {}            # Record number(s) in secondary PMI alt UR number
altURrec = {}            # Record number in secondary PMI which has the best match with the master PMI for this alt UR number
notMatched = {}            # Not matched after matchAltUR.py outputs checked
matchedPID = {}            # Matched after matchAltUR.py checked [PID from UR]
matchedUR = {}            # Matched after matchAltUR.py checked [UR from PID]
foundSecondaryRec = {}        # Master record number(s) for Secondary records
secondaryFoundPID = {}        # Secondary PID for Secondary record number
masterDetails = {}        # The master PMI details for records of interest
foundPID = {}            # Found after findUR.py checked [PID from UR]
foundUR = {}            # Found after findUR.py checked [UR from PID]
notFound = {}            # Not found after findUR.py outputs checked
recStatus = {}            # The found status
                # For matchAltUR.py this is binary encoded
                # 64 = Known match
                # 32 = Family Name match, 8 = Family Name sound match
                # 16 = Date of Birth match,
                # 4 = Firstname match, 2 = Firstname sound match,
                # 1 = Sex match
                # For findUR.py this is simple ranked encoding
                # 7 = Found Match
                # 6 = %Keys match (all 4 data items)
                # 5 = %KeysSsxS34, or %KeysSdm34 or %KeysSny34 match
                # 4 = %Keys123 match
                # 3 = %Keys124 match
                # 2 = %Keys134 match
                # 1 = %Keys234 match
                # 0 = No Match found
foundRec = {}            # The found patient record number(s) in the Master PMI file
foundSound = {}            # Master record sound match for Secondary records
extras = {}            # The additional information about sound matches
wantedMasterRec = {}        # Master PMI records of interest
secondaryRawRecNo = 0        # Record number of raw record read in from to Secondary PMI extract file
secondaryRecNo = 0        # Record no. of record read in from saved Secondary PMI
possExtensiveMatches = {}    # Possible Extensive Matches for each secondaryRecNo
possExtensiveFinds = {}        # Possible Extensive Finds for each secondaryRecNo
