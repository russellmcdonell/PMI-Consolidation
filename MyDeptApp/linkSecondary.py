'''
Subroutines to check the linkages in the cleaned up secondary PMI

NOTE: This code will need to be checked and probably edited for each different PMI in order to refect the sematics of the linkage columns.
For instance the 'Alias' value may need to be check to ensure that it is a valid number and not the UR number for another institution.
And the 'Merged' value may have four states to reflect 'newly created', 'updated', 'merged' and 'unmerged'
Similarly the 'Deleted' flag may have three values for 'newly created', 'deleted' and 'undeleted'
'''

# pylint: disable=invalid-name, bare-except

import data as d


def secondaryIsAlias():
    '''Check if the current cleaned up secondary PMI record indicates that it is an alias of another record
    Test that the 'Alias' column is not empty
    NOTE - Uses d.secondaryHas['Alias']'''
    if 'Alias' not in d.secondaryHas:
        return False
    if d.secondaryLinks['aliasLink'] is None:
        return False
    if not d.csvfields[d.secondaryIs[d.secondaryHas['Alias']]]:
        return False
    return True


def secondaryIsMerged():
    '''
Check if the current cleaned up secondary PMI record has been merged to/from another
Test that the 'Merged' column is not empty
NOTE - Uses d.secondaryMergedLink, d.secondaryHas['Merged']
    '''

    if 'Merged' not in d.secondaryHas:
        return False
    if d.secondaryLinks['mergedLink'] is None:
        return False
    if not d.csvfields[d.secondaryIs[d.secondaryHas['Merged']]]:
        return False
    return True


def secondaryIsDeleted():
    '''
Check if the current cleaned up secondary PMI record is marked as 'deleted'
Test the value in the 'deleted' flag column
NOTE - Uses d.secondaryHas['Deleted']
    '''

    if 'Deleted' not in d.secondaryHas:
        return False
    if d.csvfields[d.secondaryIs[d.secondaryHas['Deleted']]] == 'D':
        return True
    return False
