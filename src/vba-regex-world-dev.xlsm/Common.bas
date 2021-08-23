Attribute VB_Name = "Common"
Public Const E_OK = 0
Public Const E_NOT_OK = 1

' Start input cell entry
Public Const START_DATA = 11

' Collumn data
Public Enum SEARCH_COL
    SEARCH_REGEX_COL = 2
    SEARCH_FILE_COL = 3
    SEARCH_OUTPUT_COL = 4
End Enum

Public Enum REPLACE_COL
    REPLACE_REGEX_COL = 2
    REPLACE_PATTERN_COL = 3
    REPLACE_FILE_COL = 4
    REPLACE_OUTPUT_BEFORE_COL = 5
    REPLACE_OUTPUT_AFTER_COL = 6
End Enum

Public Enum FILTER_COL
    FILTER_REGEX_COL = 2
    FILTER_TYPE_COL = 3
    FILTER_FILE_COL = 5
End Enum

' Sheet data
Public Enum SHEET_INDEX
    SHEET_SEARCH = 1
    SHEET_REPLACE = 2
    SHEET_FILTER = 3
End Enum



