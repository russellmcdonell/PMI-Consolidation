
# pylint: disable=invalid-name, bare-except, line-too-long

'''
Subroutines to check the linkages in the cleaned up master PMI

NOTE: This code will need to be checked and probably edited for each different PMI in order to refect the sematics of the linkage columns.
For instance the 'Alias' value may need to be check to ensure that it is a valid number and not the UR number for another institution.
And the 'Merged' value may have four states to reflect 'newly created', 'updated', 'merged' and 'unmerged'
Similarly the 'Deleted' flag may have three values for 'newly created', 'deleted' and 'undeleted'
'''

import data as d


def masterIsAlias():
    '''
Check if the current cleaned up master PMI record indicates that it is an alias of another record
Test that the 'Alias' column is not empty
NOTE - Uses d.masterHas['Alias']
    '''
    if 'Alias' not in d.masterHas:
        return False
    if d.masterLinks['aliasLink'] is None:
        return False
    if not d.csvfields[d.masterIs[d.masterHas['Alias']]]:
        return False
    return True


def masterIsMerged():
    '''
Check if the current cleaned up master PMI record has been merged to/from another
Test that the 'Merged' column is not empty
NOTE - Uses d.masterMergedLink, d.masterHas['Merged']
    '''

    if 'Merged' not in d.masterHas:
        return False
    if d.masterLinks['mergedLink'] is None:
        return False
    if not d.csvfields[d.masterIs[d.masterHas['Merged']]]:
        return False
    return True


def masterIsDeleted():
    '''
Check if the current cleaned up master PMI record is marked as 'deleted'
Test the value in the 'deleted' flag column
NOTE - Uses d.masterHas['Deleted']
    '''

    if 'Deleted' not in d.masterHas:
        return False
    if d.csvfields[d.masterIs[d.masterHas['Deleted']]] == 'D':
        return True
    return False
