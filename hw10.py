# Update the test case automation road map with new test cases and versions downloaded from a repository
# Read in test case file names and version from a csv file
# open the road map using pandas
# Compare test cases, (file names), with what is on the Sync tab, column E , row 15 - end of sheet
# Compare the version with what is on the Sync tab, column P
# Check for duplicate test cases, (version and name will be the same)
# If duplicates found, list the duplicates on the DUP sheet with sheetname
# If the test case is on the road map and the version matches, mark column R with an x to indicate that this
#   test is in the test case text file
# If the test case is not on the road map  add it to the end of column E on the Sync sheet and add the version
#    to column O
