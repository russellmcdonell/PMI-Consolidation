# PMI-Consolidation
Consolidate two lists of names, birthdate and sex, which share a common identifier.

PMI Consolidation is the process of checking two lists of names, birthdates and sex; a master list and a secondary list,
to ensure that every name in the secondary list is also in the master list with the same birthdate and sex. The two lists will share a common identifier, so the match is "same identifier, same name, same birthdate, same sex" in both lists. Some names in the secondary list will exist in the master list, but the identifiers won't match. Here, the master list is deemed to be the source of truth and these names, in the secondary list, will have their identifier updated to the correct value. Some names in the secondary list won't exist in the master list. These names will have to be added to the master list and a master list identifier assigned, after which the matching names, in the secondary list, will be assigned the new identifier from the master list. At this point, the secondary list becomes a wholly contained subset of the master list and the two lists are consolidated.

A typical scenario, in a healthcare/hospital environment, is where a department has been running their own application for their own patients. The hospital will have a Patient Administration System (PAS) which records patient names and assigned a hospital wide UR or MRN. The department, in their independent application, will  have been creating patient records and may have been recording the hospital wide UR/MRN manually. These deparmental records will not align with the records in the PAS as they will have spelling mistakes in the names and incorrect UR/MRN value due to typos, digital transposition, etc. In this scenario, the list of all patients and UR/MRN from the PAS would be considered the master list and a list of all the department's patients and UR/MRN would be considered the secondary list. Consolidating these two lists would be useful from a data quality perspective. However, if the department intends to integrate with the PAS using HL7, then consolidating these two lists becomes essential, to avoid accidental reassignment of medical records to the incorrect patients.

Lists of doctors is another typical healthcare/hospital use case. The Pharmacy will have an application which will hold all the names of the doctors (employees) and their Prescriber number (authorization to prescribe medicines). The hospital will have a Patient Administration System (PAS) which has a list of doctors (employees) and a Provider number, for assigning doctors to inpatients. And the Medical Workforce team will have an Excel spreadsheet of all the doctors (employees) with both their Prescriber number and their Provider number. Consolidating these three lists will be essential if the hospital intends to deploy an EMR where doctors, assingned to patients, can prescribe medications.

## The Process
PMI Consolidation is not something that comes "out of the box". Each scenario will be different; each scenario will require configuration and potentially fine tuning the software. What is common is that all lists will be "dirty"; duplicates and less than perfect data (some users append an asterisk to the surname if the person is dead, others put the person's maiden name after their surname in brackets, others put nicknames in brackets after the given name). So the first two steps in the process focus on cleaning up the master and secondary lists. The next step is focuses on check records in both lists that share the same common identifier; if they have the same identifier, then are the names, date of birth and administrative sex the same. Where there is a mismatch (same identifier, different names) then we assume that the names are correct, but the identifier in the secondary list is wrong. The fourth step focuses on names in the secondary list which don't have a value for the common identifier and name in the secondary list with the wrong identifier, as discovered in third step. Here we check each of these names against every name in the master list, looking for a match.

In steps 3 and 4 we are doing matching. Hopefully everything is either a perfect match or a perfect mismatch. Alas, it never turns out that way. There will always be close, approxiamate, nearly matches; RUSSEL doesn't match RUSSELL, or does it? Often, checking additional data such as address, medicare number, next of kin etc., will clarify things. It is possible to configure these additional checks, but you can't construct a perfect, definitive matching algorith; there will always be a need for human participation in the process. To aid this manual part of the proces, steps 3 and 4 create Excel workbooks of names that need manual checking with details of the reason why and what needs checking. These steps also include tools to facilitate processing these workbooks after they have been manually checked.

### Step 1 - clean up the master list - cleanMaster.py
Cleaning the master list converts the raw list/extract into a standard form/format suitable for matching; names are cleans and capitalized, birthdates have the same format and administrative sex uses a standard codeset. The software/configuration for cleansing the master list will be different to the software/configuration for cleansing the secondary list, just as the master list will be different to the secondary list. Hence, the software (cleanMaster.py/linkMaster.py) and configuration (master.cfg) files will reside in a different folder to the software/configuration for cleansing the secondary list.

Optimizing the cleansing process involves two things; setting the correct configuration and, potentially, adjusting the Python code for any specific master link idiosyncracities. The configuration is required to tell the software how many columns of data there are in the master list and which columns contain the identifier (UR), givenName, familyName, birhtdate and sex. To facilitate the manual checking, is it possible to extract extra data that could exist in both lists and include that data in the standardized file. It won't be cleansed but it will appear in reports. Reviewing the code, and potentially amending it, is highly recommended. The code can do additional checking, such as checking that the identifier has the correct number of characters and, if it is a number, that it only contains digits. The code can also check that the identifier starts with the correct digit/character, or is in the correct range. The "out of the box" code removes things from names that are enclosed in round brackets. You may choose not to do that as it is done consistently in both the PAS and the departmental application. Or you may find that your users enclose things in angle brackets and you need to adjust the code to cater for that behaviour.

### Step 2 - clean up the secondary list - cleanSecondary.py
This step follows the same process as step one, just in a different folder. If you are consolidating more than one departmental application, then each will have it's own folder with a copy of the cleansing software (cleanSecondary.py, linkSecondary.py) and the secondary configuration (secondary.cfg). Each copy of the secondary cleansing code and configuration will need to be optimized for each departmental application extract.

**Note** Consolidation doesn't alway happen just once. Sometimes you will need to do it multiple times with a new extract each time. And sometimes the extract change in format/layout over time. To facilitate this, each master and secondary folder can have a subfolders, one for each specific extract, which will hold the specific extract plus a smaller configuration file (extract.cfg) which describes any changes in column/layout. The cleaned up data (secondary.csv) will also be saved in the matching extract subfolder.

### Step 3 - match records based up the common identifier - matchAltUR.py
matchAltUR.py check records from the cleaned up master list (master.csv) and the cleaned up secondary list (secondary.csv) where records have the same value for the identifier (UR in master.csv and AltUR in secondary.csv). matchAltUR.py will check that the names, birthdate and sex are idential for each pair of matching records. matchAltUR.py outputs lists of perfect matches, secondary records that have no matching identifier in the master list and Excel workbooks of records where the identifiers match, but something else didn't (same name with different birthdates, etc.) - the ambiguous matches. These ambiguous matches will need to be checked manually and tagged as either an actual match match (typo in the birthdate) or an actual non-match (father and son with the same name, different birthdates). Things that are marked as non-matches are still included in the process; they will be checked against all records in the master list for match of just name, birthdate and sex, on the assumption that those details in the secondary list are correct, but the identifier is wrong.

### Step 4 - search for name/birthdate/sex matches for unmatched records - findUR.py
findUR.py uses the outputs of step 3 and the manual checking to identify secondary list records that remain unmatched. findUR.py checks each unmatched record against all the records in the master list, looking for a name/birthdate/sex match. findUR.py outputs a list of perfect matches (finds), secondary records where nothing matches anything in the master list (notFound) and Excel workbooks of records where parts of name/birthdate/sex match and some part don't - the ambiguous finds, which will include the secondary list record and all possible master list records that might be a match. These ambiguous finds will need to be checked manually and tagged as either an actual match match (multiple typos) or an actual non-match/new patient. Once this has been done there will be tfour lists.
* Records from the secondary list that have an exact match in the master list (from step 2).
* Records from the secondary list that have the correct identifier, but have typos in name and/or birthdate and/or sex (from step 2 - the correct name/birthdate/sex will be listed and these records will need to be updated).
* Records from the secondary list that are a perfect match with a master record based upon name/birthdate/sex, but they have no identifier or the wrong identifier (from step 3 - the correct identier has been found and the secondary record will been to be updated).
* Records from the secondary list which do not exist in the master list (from step 3 - these records will need to be added to the master list, an identifier will have to be allocated and secondary record will need to be updated to assign this identifier).

Full descriptions of the theory and process can be found in the documentation in the Documentation folder.

