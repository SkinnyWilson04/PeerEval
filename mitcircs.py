"""
Updated version 2.2.2 of the Mitigating Circumstances program.
Rowan Lawrence - new as of 13/01/2025
"""
import os
import re
import pandas
from datetime import datetime
from datetime import date
from typing import TypeVar, List, Dict
from dataclasses import dataclass
import tkinter as tk
from tkinter import filedialog
from tkinter.messagebox import showinfo


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


Arguments = TypeVar('argparse.Namespace')
Array     = TypeVar('numpy.ndarray')
Window    = TypeVar('tkinter.Tk')
DataFrame = TypeVar('pandas.core.frame.DataFrame')
Series    = TypeVar('pandas.core.series.Series')

MINIMUM_REQUIRED_RESPONSES = 4


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Each assessment that can be applied for in the Qualtrics form has a fixed set of subquestions
# relating to the Year, Unit Code, whether it is a Resubmission, etc. These are suffixed by the
# same strings which are assigned to variables here to make this more descriptive and easier to
# modify in the future, in case the column-names in the Qualtrics form change.
# "_SUFFIX" indicates that these strings are suffixes to the column names, so must be located
# with the .endswith() method ONLY, or should be joined to the *end* of a column name
# e.g., the name of the column containing the unit code for a response 3 would be constructed:
#  >>> question_num: str = "3"
#  >>> colname: str = f"{question_num}_Q161{COLNAME_SUFFIX_UNITCODE}"

COLNAME_SUFFIX_DIVISION         = "_1"         # Name of the Division the student belongs to
COLNAME_SUFFIX_PROGRAMME        = "_2"         # The Programme the student is currently studying on
COLNAME_SUFFIX_COURSEYEAR       = "_3"         # Current Year within that course, as an Integer (1, 2, 3, etc.)
COLNAME_SUFFIX_UNITASSESSMENT   = "_4"         # Unit Title and Name of the Assessment being applied for
COLNAME_SUFFIX_OTHERINFORMATION = "_Q37"       # "Other" Assessment info: Submission date information in format: <Unit Code>: <Assessment Name> - <Submission Date>
COLNAME_SUFFIX_RESUBMISSION     = "_Q165"      # Is this a Resubmission? Date is provided if it is (otherwise "No")
COLNAME_SUFFIX_RESUB_FIRST      = "_1_TEXT"    # Resubmission deadline *if* this is a 1st attempt
COLNAME_SUFFIX_RESUB_SECOND     = "_3_TEXT"    # Resubmission deadline *if* this is a 2nd attempt
COLNAME_SUFFIX_SUBSTATUS        = "_Q163"      # Submission status and intentions - have you submitted the work, attended exam, etc? Will you?

# OLD_COLNAME_SUFFIX_PROGRAMME    = "_1"         # Name of degree programme the student is currently on 
# OLD_COLNAME_SUFFIX_COURSEYEAR   = "_2"         # Integer, Current Year within that course (1, 2, 3, etc.)
# OLD_COLNAME_SUFFIX_UNITCODE     = "_3"         # Unit code for the Module the Assessment being applied for is from
# OLD_COLNAME_SUFFIX_ASSESSMENT   = "_4"         # Name (and / or submission date) of the Assessment
# OLD_COLNAME_SUFFIX_RESUBMISSION = "_Q1"        # Is this a Resubmission? Date is provided if it is (otherwise "No")
# OLD_COLNAME_SUFFIX_RESUB_FIRST  = "_1_TEXT"    # Resubmission deadline *if* this is a 1st attempt
# OLD_COLNAME_SUFFIX_RESUB_SECOND = "_4_TEXT"    # Resubmission deadline *if* this is a 2nd attempt
# OLD_COLNAME_SUFFIX_SUBSTATUS    = "_Q165"      # Submission status - have you submitted the work, attended exam, etc?


# - - - - - - >
# Alternatively, since the search procedure for the response columns in each row returns a series,
# we could just look up the index of each response column from a dict directly... Bit more ergonomic?

RESPONSE_COLUMN_INDICES: Dict[str, int] = {COLNAME_SUFFIX_DIVISION:         0, COLNAME_SUFFIX_PROGRAMME:      1,
                                           COLNAME_SUFFIX_COURSEYEAR:       2, COLNAME_SUFFIX_UNITASSESSMENT: 3,
                                           COLNAME_SUFFIX_OTHERINFORMATION: 4, COLNAME_SUFFIX_RESUBMISSION:   5,
                                           COLNAME_SUFFIX_RESUB_FIRST:      6, COLNAME_SUFFIX_RESUB_SECOND:   7,
                                           COLNAME_SUFFIX_SUBSTATUS:        8}

# NOTE:
# Another artefact of the previous 2024 Qualtrics version that we were using as input.
# The same is true for the COLNAME_PREFIX global variables defined above.
# The newest (NOVEMBER 2024) version of the Qualtrics output has had it's Q-columns rearranaged and,
# it seems, had new ones added to it. Again, these are kept here in case of another change / reversion.
#OLD_RESPONSE_COLUMN_INDICES: Dict[str, int] = {COLNAME_SUFFIX_PROGRAMME:    0, COLNAME_SUFFIX_COURSEYEAR:  1,
#                                               COLNAME_SUFFIX_UNITCODE:     2, COLNAME_SUFFIX_ASSESSMENT:  3,
#                                               COLNAME_SUFFIX_RESUBMISSION: 4, COLNAME_SUFFIX_RESUB_FIRST: 5,
#                                               COLNAME_SUFFIX_RESUB_SECOND: 6, COLNAME_SUFFIX_SUBSTATUS:   7}


# - - - - - - >
# As above, there are common columns on the Qualtrics form where students may enter identifying
# and supplementary information, such as their name and email address, their academic advisor,
# reason for mitigation, etc. As above, these are assigned to variables here to improve clarity
# and give a common point to change in case these need to be adjusted in future versions.
# "_PREFIX" indicates that these strings are prefixes to the column name (or otherwise constitute
# the entirety of the target column name - they function the same either way)
# NOTE: The column name "Q1" is shared by both the Student Name and Supervisor Name columns,
#       oddly - convenient-ish is the behaviour of [] in that df['Q1'] returns the first matching
#       column, which  is always the Student Name. This does, however, mean that more parsing
#       is required to get at the Supervisor Name, hence the reminder here to use the information
#       in the top "Header" row of the sheet instead

COLNAME_PREFIX_DATESUBMITTED    = "RecordedDate" # String - Date and time the application was submitted
COLNAME_PREFIX_STUDENTNAME      = "Q1"           # Full name of the student applying
COLNAME_PREFIX_EMAILADDRESS     = "Q3"           # UoM (I assume) email address of the student applying
COLNAME_PREFIX_STUDENTID        = "Q4"           # Student ID number of applicant (as on student card)
COLNAME_PREFIX_ISPOSTGRADORRES  = "Q150"         # Is this a Postgrad. dissertation or Research project?
COLNAME_PREFIX_SUPERVISCONTACT  = "Q152"         # If Yes, have you spoken to your dissertation supervisor?
COLNAME_PREFIX_SUPERVISORNAME   = "UseHeader"    # "Q1" CLASH - Name of dissertation supervisor
COLNAME_PREFIX_TIER4_VISA       = "Q153"         # Are you on a Tier 4 Visa?
COLNAME_PREFIX_PROPOSEDDEADLINE = "Q151"         # Proposed new deadline, as agreed by project supervisor
COLNAME_PREFIX_ADVISORNAME      = "Q17"          # Name of your academic advisor
COLNAME_PREFIX_MITIGATIONDETAIL = "Q19"          # Summary of circumstances requiring mitiating circumstances
COLNAME_PREFIX_PERIODAFFECTED   = "Q20"          # dd/MM/YY to dd/MM/YY period affected
COLNAME_PREFIX_LATEAPPLICATION  = "Q21"          # Appl. must be within 5 days of sub. date - if late, why?
COLNAME_PREFIX_ASSESSMENTCOUNT  = "Q160"         # How many assessments are you applying for?
COLNAME_PREFIX_DASS_REGISTERED  = "Q2"           # Are you currently Disability Advisory Support Service registered?
COLNAME_PREFIX_PROVIDESEVIDENCE = "Q3"           # Will you be providing evidence along with your application? (NOTE: Clash!)
COLNAME_PREFIX_EVIDENCEFILENAME = "Q164_Name"    # If "Yes", the filename of the evidence submitted

# INFORMATION:
# Column name Q160, containing the specified assessment application count, *used* to be 
# column name Q163 in the older test data. Making a note of this here in case the format 
# changes again or reverts in the future.
#COLNAME_PREFIX_ASSESSMENTCOUNT  = "Q163"         # How many assessments are you applying for?


# - - - - - - >


COLUMN_KEY_ERROR_MESSAGES: Dict[int, str] = {
    COLNAME_PREFIX_STUDENTID:        "Could not read 'Student ID' column!\nExpected column name to be 'Q4' - please check this and run again.",
    COLNAME_PREFIX_EMAILADDRESS:     "Could not read 'Email Address' column!\n Expected column name to be 'Q3' - please check this and run again.",
    COLNAME_PREFIX_DATESUBMITTED:    "Could not read '' column!\n Expected column name to be '' - please check this and run again.",
    COLNAME_PREFIX_ISPOSTGRADORRES:  "Could not read 'Is Postgrad. or Research' column!\n Expected column name to be 'Q150' - please check this and run again.",
    COLNAME_PREFIX_ASSESSMENTCOUNT:  "Could not read 'Assessment Count' column!\n Expected column name to be 'Q160' - please check this and run again.",
    COLNAME_PREFIX_MITIGATIONDETAIL: "Could not read 'Details of Mitigation' column!\n Expected column name to be 'Q19' - please check this and run again.",
    COLNAME_PREFIX_PERIODAFFECTED:   "Could not read 'Period Affected' column!\n Expected column name to be 'Q20' - please check this and run again.",
    COLNAME_PREFIX_ADVISORNAME:      "Could not read 'Advisor Name' column!\n Expected column name to be 'Q17' - please check this and run again.",
    COLNAME_PREFIX_LATEAPPLICATION:  "Could not read 'Application Outside of Deadline' column!\n Expected column name to be 'Q21' - please check this and run again.",
    COLNAME_PREFIX_SUPERVISCONTACT:  "Could not read 'Supervisor Spoken To' column!\n Expected column name to be 'Q152' - please check this and run again.",
    COLNAME_PREFIX_TIER4_VISA:       "Could not read 'On Tier 4 Visa?' column!\n Expected column name to be 'Q153' - please check this and run again.",
    COLNAME_PREFIX_PROPOSEDDEADLINE: "Could not read 'Proposed New Submission Date' column!\n Expected column name to be 'Q151' - please check this and run again."
}


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


SEPARATOR: str = "\n - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - \n"
LOGLINE:   int = 0


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# If the input string / cell is empty or NaN, return a helpful string
# Otherwise, simply return the input string

def string_reformat_nan(cell_in: str, empty_string: str = "None given") -> str:
    if pandas.isna(cell_in) or cell_in == "" or cell_in.lower() == "nan" or cell_in == None:
        return empty_string
    else:
        return cell_in


# - - - - - - >
# Convert bool value to a string
# (Simply doing str(b) might also work?)

def string_reformat_bool(b: bool) -> str:
    if b:
        return "True"
    else:
        return "False"


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Mitigating Circumstances Request class
# NOTE: That instances of StudentRequest are per student, *not* per assessment.
#       If students apply for mitigating circumstances for multiple assessments, then each
#       assessment will have its own entries in the Lists prefixed with "asm_"
# NOTE: The output tracker also requires a "Division", but this isn't in the input raw data -
#       the field here is a placeholder and will be left empty for the moment. Presumably this
#       will be added to the Qualtrics form in the near future
class StudentRequest:
    def __init__(self):
        self.name:            str = ""          # Name of the student applying
        self.ID:              str = ""          # University ID of student
        self.email:           str = ""          # University email of student
        self.subdate:         str = ""          # Date on which the Qualtrics survey was submitted
        self.programme:       str = ""          # Programme student is on
        self.courseyear:      str = ""          # Year on their Programme the student is currently in
        self.division:        str = ""          # Division the Programme belongs to
        self.isPGR:           str = ""          # Whether the application is for a Postgraduate Dissertation or Research project
        self.NAffected:       str = ""          # The number of Assessments that the student states have been affected
        self.circumstances:   str = ""          # The circumstances requiring mitigating circumstances
        self.dates_affected:  str = ""          # Period affected by the above circumstances
        self.DASS:            str = False       # Whether the student is DASS registered
        self.semester:        str = ""          # The semester that has been affected
        self.advisor:         str = ""          # The Academic Advisor of the student
        self.latereason:      str = ""          # If this application is being submitted late, this is the reason for that late sub.
        self.evidence:        str = ""          # Whether there is supporting evidence submitted with the application
        self.evidencesummary: str = ""          # A brief summary of the evidence, if submitted
        self.superinformed:   str = ""          # Whether the student has informed the supervisor of their circumstances
        self.supervisor:      str = ""          # The name of the Supervisor
        self.T4Visa:          str = ""          # Whether this is an overseas student on a Tier 4 Visa
        self.proposedDL:      str = ""          # The new proposed submission deadline
        self.asm_codes:       List[str] = []    # Unit Codes of the Assessments affected
        self.asm_names:       List[str] = []    # Names of the Assessments affected
        self.other_asm:       List[str] = []    # If "Other" Assessments are selected, this is the information entered by the student
        self.asm_is_resub:    List[str] = []    # Whether this is a re-submission (rather than simply a late submission)
        self.asm_resubdate:   List[str] = []    # If it is a Resubmission, this is the new Resubmission Deadline
        self.asm_resubstatus: List[str] = []    # Whether or not they have submitted the work, attended the exam, etc.


    def to_string(self) -> str:
        req: str = f"\n\nRequest by Student '{self.name}' (ID: {self.ID}, DASS: {self.DASS}) on Year {self.courseyear} of Programme '{self.programme}':"
        req = f"{req}\n  > With Advisor '{self.advisor}' and Supervisor '{self.supervisor}'"
        req = f"{req}\n  > Uni. Email:   {self.email}"
        req = f"{req}\n  > Tier 4 Visa?  {self.T4Visa}"
        req = f"{req}\n  > Submitted on: {self.subdate}"
        req = f"{req}\n\nDates affected: {self.dates_affected}, due to:\n  > '{self.circumstances}'\n\nStudent applied for the following assessments:\n"
        
        for assessment in range(len(self.asm_codes)):
            req = req + f"{assessment + 1}. {self.asm_names[assessment]}\n  ? Unit Code:     {self.asm_codes[assessment]}\n  ? Resubmission:  {self.asm_is_resub[assessment]}"
            req = req + f"\n  ? Proposed Date: {self.asm_resubdate[assessment]}\n  ? Resub. Status: {self.asm_resubstatus[assessment]}\n\n"

        req = f"{req}{SEPARATOR}\n"
        return req
    

    def __str__(self):
        return str(self.to_string())


# - - - - - - >


@dataclass
class Counter:
    counter: int

    def __init__(self):
        self.counter = 1


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Given an column of input cells in a dataframe which contain one or more newline-separated
# strings, identify the maximum character-width of the column required to accomodate the longest
# string. Basically - if we were to display all of the strings in each cell on new lines, how
# wide would the resulting column need to be to ensure that none of the strings are cut off?

def required_column_width(dataframe: DataFrame, column: str) -> int:
    cell_max: int = 0
    this_max: str = 0
    
    for cell in dataframe[column]:
        lines: List[str] = cell.split('\n')
        this_max = max(lines, key = len)
        if len(this_max) > cell_max:
            cell_max = len(this_max)

    print(f"{column} required width = {cell_max} (string: '{this_max}')")
    return cell_max


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Locate the index of a row in a dataframe containing some string.
# If no row containing the target string can be found, this function returns None.
# This is checked for in the drop_row() function and the entire row is dropped in place if found.

def row_containing(dataframe: DataFrame, string: str) -> int:
    location: int = None

    for index, row in dataframe.iterrows():
        values: List[str] = [value for value in row]
        try:
            contains: List[bool] = [True for value in values if string in value]
            if any(contains):
                location = index
        except:
            continue
    
    return location


# - - - - - - >


def drop_row_by_string(dataframe: DataFrame, string: str) -> DataFrame:
    index: int = row_containing(dataframe, string)
    
    if not index:
        return dataframe
    
    dataframe.drop(index, inplace = True)
    return dataframe.reset_index(drop = True)


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


def max_string_length(strings: List[str]) -> int:
    string: str = max(strings, key = len)
    return len(string)


# - - - - - - >


def create_output_filename(directory: str, N_applications: int) -> bool:
   filename: str = f"Mitigating Circumstances Tracker - {N_applications} Students - {date_today()}.xlsx"
   return os.path.join(directory, filename)


# - - - - - - >


def display_list(lst: List[any]):
    for obj in lst:
        print(obj)


# - - - - - - >


def display_dict(dct: Dict[any, any]):
    for key in dct.keys():
        print(f"{key} => {dct[key]}")


# - - - - - - >


def log_string(logfile: str, string: str, logcount: Counter) -> None:
    string = f"Log {logcount.counter} > {string}\n"
    with open(logfile, 'a') as fptr:
        fptr.write(string)
    logcount.counter = logcount.counter + 1


# - - - - - - >


def log_request(logfile: str, request: StudentRequest, logcount: Counter):
    request_string: str = f"\nLog {logcount.counter}\nRequest instance at location: {hex(id(request))}\n{request.to_string()}"

    with open(logfile, 'a') as fptr:
        fptr.write(request_string)


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


def date_today() -> str:
    today_string: str = date.today().strftime("%d-%B-%Y")
    return today_string


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Helper functions to get the current time and date as strings
# Values are separated by hyphen '-' to allow for use in filenames - otherwise, use the
# 'current_datetime()' function for better formatting and split on the space as needed

def current_time() -> str:
    time_string: str = datetime.now().strftime("%H-%M-%S")
    return time_string


# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ >


def current_date() -> str:
    date_string: str = datetime.now().strftime("%d-%m-%Y")
    return date_string


# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ >


def current_datetime() -> str:
    datetime_string: str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    return datetime_string


# ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ >


def object_exists(path: str, suppress: bool) -> bool:
    try:
        assert os.path.exists(path) == True
        return True
    except AssertionError:
        if not suppress:
            print(f"Error: Cannot find file or directory '{path}'\n    -> Check filepath for this object and try again.")
        return False


# - - - - - - >


def is_filetype(filepath: str, extension: str) -> bool:
    return os.path.splitext(filepath) == extension


# - - - - - - >


def create_file(filepath: str):
    with open(filepath, 'w') as _:
        pass


# - - - - - - >


def create_log_if_requested(directory: str, logging: bool) -> str:
    if not logging:
        return None
    else:
        logpath: str = os.path.join(directory, f"MitCircLog_{current_time()}_{current_date()}.txt")
        print(f"Creating file: '{logpath}'...")
        create_file(logpath)
        return logpath


# - - - - - - >


def find_columns(colnames: List[str], pattern: str) -> List[str]:
    pat = re.compile(pattern)
    names: List[str] = [name for name in colnames if pat.match(name)]
    return names


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# In the updated data, students are required to enter the Unit Code, Assessment Name, and
# Submission Date all into one cell. In the output, we need to write the Unit Code separately.
# Great stuff.
# Of course, students could not be trusted to do this properly so we cannot assume that the
# input is well-formed. The format *should* be a small number of UPPERCASE letters, followed by
# a small number of digits. This is separated from the rest of the cell content by a colon ':'.
# First, strip any leading or trailing whitespace from the input string.
# Second, use a regex to detect and validate the Unit Code.
# If a substring matching this Unit Code format is provided, then this is simply returned.
# Otherwise, a placeholder is created from the first 9 characters of the input - since it's
# *possible* that they have provided something sensible here but simply in the wrong format,
# we want to try and include this information on the output spreadsheet rather than binning it.

def detect_return_unitcode(string: str) -> str:
    string = string.strip()
    pattern = re.compile(r"""
    ^[A-Z]{1,5}    # Starts with between 1 and 5 uppercase letters
    [0-9]{2,6}     # Immediately followed by between 2 and 6 digits
    """)

    result = re.search(pattern, string)

    if not result:
        return f"{string[:9]}"
    else:
        return result.group(0)


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Return a dict of integer lists of all of the column indices relating to each unique assessment response.
# For example, let's say that a student applies for the 3rd-year Nursing assessment "NURS34555", which
# occupies response columns beginning with "6_Q*". These response columns were identified by the
# 'locate_response_columns()' function below. Here, this function identifies that the relevant columns
# are located between indices "64" - "71" inclusive, converts these strings to integers, and returns
# them as a list of column indices.
# e.g.:
#    ...
#    number = "6_Q"
#    ...      # Work to parse the "6" from the column name
#    responses[number] = [64, 65, 66, 67, 68, 69, 70, 71]
#    ...      # Do the rest of em
#    return responses

def unique_response_locations(response_columns: Dict[str, int], response_numbers: int, logging: bool, logfile: str, logcount: Counter) -> Dict[str, List[int]]:
    responses: Dict[str, List[int]] = {}
    column_names: List[str] = response_columns.keys()
    for response in response_numbers:
        target_cols: List[str] = [col for col in column_names if col.startswith(f"{response}_")]
        indices: List[int] = [int(response_columns[col]) for col in target_cols]
        if indices:
            responses[response] = indices
    
    if logging:
        for resp in responses.keys():
            log_string(logfile, f"Response: '{resp}'\n  > Indices:  {responses[resp]}", logcount)
    
    return responses


# - - - - - - >
# The task here is to systematically identify the columns which contain relevant response data for each
# module and assessment, and to return these in a dict containing the name and the index of that column.
# An additional function, 'structure_response_columns()' (below) will do the actual work of identifying
# columns that have responses in them and structuring the relevant information.
# So, these two functions are effectively the heart of identifying and restructuring the relevant
# information for each assessment that a student selects.


def locate_response_columns(dataframe: DataFrame, display_index: bool, response_max: int, logging: bool, logfile: str, logcount: Counter) -> Dict[str, List[int]]:
    # All of the relevant response columns for each assessment start with a number
    # (at the moment the maximum is 30, though I assume this will grow as more
    # assessments are added so I've left some headroom here with range of 1-99).
    # NOTE: The response columns are the ONLY ONES that follow this "^[num]_Q1*"
    #       format for their column names, so I'm making the assumption that
    #       this WILL NOT CHANGE in future versions of the spreadsheet. Otherwise,
    #       it will become much more difficult (potentially impossible) to
    #       correctly identify and parse the relevant columns! No touch! Bad!
    numbers = tuple([str(number) for number in range(1, response_max)])
    targets = dataframe.columns[dataframe.columns.str.startswith(numbers)]
    indices = dataframe.columns.get_indexer(targets)
    columns = {}

    if logging:
        log_string(logfile, "Locating assessment response columns in dataframe:", logcount)

    for name, index in zip(targets, indices):
        columns[name] = index

        if display_index:
            print(f"  > Response Column: {name}   At: {index}")
        if logging:
            log_string(logfile, f" Response Column: {name}   At: {index}", logcount)

    locations: Dict[str, List[int]] = unique_response_locations(columns, numbers, logging, logfile, logcount)

    return locations


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


def extract_top_row(dataframe: DataFrame) -> DataFrame:
    return dataframe.head(1)


# - - - - - - >


def delete_top_row(dataframe: DataFrame) -> DataFrame:
    return dataframe.tail(-1)


# - - - - - - >
# In the unusual event that a student does not give both a first and last name, we're going to handle building
# it a little more thoroughly. First, check if the corresponding cell for each name is empty. If not, then
# strip off the whitespace and assign to both firstname and lastname, respectively. Join together the two
# with a space in-between and return.
# Note that the additional call to .strip() on fullname is necessary to prevent a leading or trailing space
# should the student fail to provide a firstname or surname, respectively (again, probs won't happen but still)

def create_student_name(row: DataFrame) -> str:
    fullname: str = ""

    if not pandas.isna(row['Q1']):
        fullname = str(row['Q1'])

    return fullname.strip()


# - - - - - - >


def student_is_DASS(row: DataFrame, return_input_string: str) -> bool | str:
    cell = row[COLNAME_PREFIX_DASS_REGISTERED]
    if cell == "" or cell == 0 or pandas.isna(cell):
        return False
    else:
        if return_input_string:
            return str(cell)
        return True


# - - - - - - >
# The "header" is the top row (NOT the column names!) of the input Excel sheet. It contains
# the full string of each question students were asked, and so provides some supplementary
# information to the responses provided in each column.
# This is used here to identify both the Programme and Year that the student is on from the
# assessment(s) that they have applied for. This requires some massaging of the dataframe,
# and is done by:
#   1. Pulling all of the strings in the header out into a list
#   2. Identifying the indices of columns which specify "Programme" in their header
#   3. Subsetting the row to only those "Programme" columns
#   4. Removing any cells which are empty / NaN
#   5. Casting the value to a string and passing this to the "reformat if nan" function,
#      which will replace this with a more helpful "None Given" string if this is missing
#      or there is some other issue with it
#   6. Returning the contents of the first "Programme" column in the row as a string;
#      given that the columns have been carefully selected and cleaned, this should
#      contain their programme name
# NOTE: In the test Excel sheet, "students" have made applications for multiple assessments,
#       some of which are on different Programmes. I'm not sure how often this will occur in
#       reality, but if so this logic will need to be changed

def string_parse_header(row: DataFrame, header: DataFrame, searchstring: str, logging: bool, logfile: str, logcount: Counter) -> str:
    strings: List[str] = header.iloc[0, :].astype(str).tolist()
    indices: List[int] = [index for index in range(len(strings)) if searchstring in strings[index]]
    years: DataFrame = row.iloc[indices]
    years = years.dropna(how = "all")
    
    try:
        return string_reformat_nan(str(years.iloc[0]))
    except IndexError as IdxErr:
        if logging:
            log_string(logfile, str(IdxErr), logcount)
        return "..."


# - - - - - - >
# NOTE: This is a potential buzzy bug-zone

def string_parse_division(row: Series) -> str:
    if not row.iloc[0] or pandas.isna(row.iloc[0]):
        return "None provided"
    return str(row.iloc[0])


# - - - - - - >
# Each request is formed of 8 response columns detailing the Unit Code of the assessment, the Assessment
# Name, whether it is a Resubmission, etc...
# This function ensures that at least the minimum number of these responses have been filled in by the
# student; if they haven't, then we can safely assume that either:
#    1. The student isn't applying for this assessment at all
#    2. They have omitted a sufficient number of details that this can be ignored
#
# In the second case, it would be extremely helpful to log that information as it is possible
# that they were in fact attempting to apply for this assessment, but there has been some other error or issue.
# NOTE: In the (unlikely) event that the minimum number of responses is longer than the number of possible
#       responses per assessment, the minimum will simply be set to the total number of responses. This means
#       that such a mistake will require *all* of the response columns to be filled in, otherwise the
#       assessment will not be added to the StudentRequest. A warning will be printed to prompt a check for this.

def minimum_responses_provided(cells: Series, minimum: int) -> bool:
    N: int = len(cells)

    if minimum > N:
        print(f"! Warning !   Length of Series ({N}) is less than the required minimum responses ({minimum})!")
        print("    > How have you managed that?")
        minimum = N

    omitted:  int = sum([pandas.isna(cell) for cell in cells])
    provided: int = max(0, (N - omitted))
    
    return (provided >= minimum)


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


def build_student_requests(qualtrics: DataFrame, display: bool, logging: bool, logfile: str, logcount: Counter) -> List[StudentRequest]:
    response_min: int = MINIMUM_REQUIRED_RESPONSES
    requests:  List[StudentRequest] = []
    header:    DataFrame = extract_top_row(qualtrics)
    qualtrics: DataFrame = delete_top_row(qualtrics)
    Q_cols:    Dict[str, List[int]] = locate_response_columns(qualtrics, display_index = True,
                                                              response_max = 100,
                                                              logging = logging, logfile = logfile, logcount = logcount)

    # Begin looping over each row in the Qualtrics data, with each row containing the submission
    # of one student. First, error-check and read some essential data including Name, ID, Email,
    # Date of Submission, Affected Dates, Via Type etc... If any of these are missing or their
    # columns are named incorrectly, then provide a clear error message from the
    # COLUMN_KEY_ERROR_MESSAGES dict, log it, and exit the program, since these are needed.
    # Otherwise, pull out all of their Q question response columns. Start by getting the Division
    # from the first of these - since the Division is in all of them (for some reason), we only
    # want to do this on the first iteration. Ensure that a given Q has actually been filled in
    # by checking that there are a sufficient number of cells which are filled in / not empty.
    # If we are below this required minimum then we can safely say this Q is not one of the ones
    # that has been filled in, so we can continue to the next *_Q

    for _, row in qualtrics.iterrows():
        try:
            req: StudentRequest = StudentRequest()
            req.name            = create_student_name(row)
            req.ID              = string_reformat_nan(str(row[COLNAME_PREFIX_STUDENTID]).strip())
            req.email           = string_reformat_nan(str(row[COLNAME_PREFIX_EMAILADDRESS]).strip())
            req.subdate         = string_reformat_nan(str(row[COLNAME_PREFIX_DATESUBMITTED]).strip())
            req.programme       = string_parse_header(row, header, "Programme", logging, logfile, logcount)
            req.courseyear      = string_parse_header(row, header, "Year", logging, logfile, logcount)
            req.isPGR           = string_reformat_nan(str(row[COLNAME_PREFIX_ISPOSTGRADORRES]).strip())
            req.NAffected       = string_reformat_nan(str(row[COLNAME_PREFIX_ASSESSMENTCOUNT]).strip())
            req.circumstances   = string_reformat_nan(str(row[COLNAME_PREFIX_MITIGATIONDETAIL]).strip())
            req.dates_affected  = string_reformat_nan(str(row[COLNAME_PREFIX_PERIODAFFECTED]).strip())
            req.DASS            = str(student_is_DASS(row, return_input_string = True))
            req.semester        = "..."
            req.advisor         = string_reformat_nan(str(row[COLNAME_PREFIX_ADVISORNAME]).strip())
            req.latereason      = string_reformat_nan(str(row[COLNAME_PREFIX_LATEAPPLICATION]).strip())
            req.evidence        = string_parse_header(row, header, "submitting evidence with your application", logging, logfile, logcount)
            req.superinformed   = string_reformat_nan(str(row[COLNAME_PREFIX_SUPERVISCONTACT]).strip())
            req.supervisor      = string_parse_header(row, header, "Dissertation supervisor name", logging, logfile, logcount)
            req.T4Visa          = string_reformat_nan(str(row[COLNAME_PREFIX_TIER4_VISA]).strip())
            req.proposedDL      = string_reformat_nan(str(row[COLNAME_PREFIX_PROPOSEDDEADLINE]).strip())
        except KeyError as kerr:
            check = COLUMN_KEY_ERROR_MESSAGES[str(kerr).replace("'", "")]
            if logging:
                log_string(logfile, check, logcount)
            tk.messagebox.showinfo(title = "Column Name Error...", message = check)
            destroy_window()
            exit(1)

        # Once we're done with the above, we need to move on to the multiple-choice questions. If the
        # corresponding cell for one of these contains NaN, then the student has not selected it and
        # so we can simply continue.
        # Additionally, since the students' Division is stored in these Q columns, we need to grab
        # this here and add copy it into the req instance once we're finished.
        division: str = ""
        first_iteration: bool = True

        for key in Q_cols.keys():
            
            # Grab all of the columns associated with a given response (e.g., "1_*" for response #1)
            # Since this is 1D, they are assigned to a simpler Series rather than a DataFrame
            cells: Series = row.iloc[Q_cols[key]]

            # Confirm that the required number of columns have been filled in for this response.
            # If the majority of the columns are unfilled, we can safely assume that the student
            # is not applying for this assessment, and it can be ignored.
            if not minimum_responses_provided(cells, response_min):
                continue

            # Next, try to grab the Division name from the first of the response cells [0]
            # This is called *after* we have confirmed that this Q has actually been filled
            # in, so the first_iteration flag does not necessarily correspond to 1_Q*
            # (not that this actually matters at all - just a note to myself)
            if first_iteration:
                division = string_parse_division(cells)
                first_iteration = False

            # Otherwise, begin assigning their responses to the relevant lists within their Request
            # instance - string_reformat_nan() ensures that a helpful string replaces any missing
            # or empty cells. Honestly, these are not as bad as they look:
            #  1. Look up the index for the required column from the RESPONSE_COLUMN_INDICES dict
            #  2. Use .iloc[index] to access that value from that cell in the row
            #  3. Cast the value to a string
            #  4. Pass that to string_reformat_nan(), which will do some sanity-checking as described above
            #  5. *The Unit Code is located from the information provided using a separate detect_return_unitcode() function
            #  6. Append the error-checked return value to the relevant list
            req.asm_codes.append(       detect_return_unitcode( str( cells.iloc[RESPONSE_COLUMN_INDICES[ COLNAME_SUFFIX_UNITASSESSMENT ]] )))
            req.asm_names.append(       string_reformat_nan(    str( cells.iloc[RESPONSE_COLUMN_INDICES[ COLNAME_SUFFIX_UNITASSESSMENT ]] )))
            req.asm_is_resub.append(    string_reformat_nan(    str( cells.iloc[RESPONSE_COLUMN_INDICES[  COLNAME_SUFFIX_RESUBMISSION  ]] )))
            req.other_asm.append(       string_reformat_nan(    str( cells.iloc[RESPONSE_COLUMN_INDICES[COLNAME_SUFFIX_OTHERINFORMATION]] ), empty_string = "-"))
            req.asm_resubdate.append(   string_reformat_nan(    str( cells.iloc[RESPONSE_COLUMN_INDICES[  COLNAME_SUFFIX_RESUB_FIRST   ]] )))
            req.asm_resubstatus.append( string_reformat_nan(    str( cells.iloc[RESPONSE_COLUMN_INDICES[   COLNAME_SUFFIX_SUBSTATUS    ]] )))
        
        # If the user checked the "Display output while running" box in the GUI, then print the full request
        # to the console here for debugging / sanity-checking. Add the Division name to the request, and then
        # finally add the completed request to the request list.
        if display:
            print(req)
        if logging:
            log_request(logfile, req, logcount)

        req.division = division
        requests.append(req)
    
    return requests


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# TODO: This entire function could be replaced by:
#
# pandas.DataFrame.from_dict(df, orient = "index", columns = colnames)
#
# ...Which would also almost certainly be faster and less error-prone

def requests_to_spreadsheet(requests: List[StudentRequest], output: str):
    # DataFrame will be formed of a Dict[str, List[...]] where the key is the column name
    # and the List is the contents of the column. Each column is here represented by a list
    subdates:      List[str] = []
    names:         List[str] = []
    emails:        List[str] = []
    IDs:           List[str] = []
    PGRs:          List[str] = []
    N_affected:    List[str] = []
    divisions:     List[str] = []
    programmes:    List[str] = []
    years:         List[str] = []
    asmcodes:      List[str] = []
    asmnames:      List[str] = []
    otherinfo:     List[str] = []
    isresubs:      List[str] = []
    resubdates:    List[str] = []
    substats:      List[str] = []
    advisors:      List[str] = []
    circumstances: List[str] = []
    periods:       List[str] = []
    latereasons:   List[str] = []
    isDASS:        List[str] = []
    evidences:     List[str] = []
    toldsupervi:   List[str] = []
    supervisors:   List[str] = []
    onT4Visas:     List[str] = []
    proposedDL:    List[str] = []
    evidencesumm:  List[str] = []
    outcomes:      List[str] = []
    emailtypes:    List[str] = []
    sendbyperson:  List[str] = []
    notes:         List[str] = []
    panelnotes:    List[str] = []
    outcomesent:   List[str] = []
    outcomedate:   List[str] = []

    # Loop over the Student Requests and build the column lists. The columns related to
    # outcomes or other data filled in during the review process are indicated by "Pending"
    # or ellipsis
    for req in requests:
        subdates.append(req.subdate)
        names.append(req.name)
        emails.append(req.email)
        IDs.append(req.ID)
        PGRs.append(req.isPGR)
        N_affected.append(req.NAffected)
        divisions.append(req.division)
        programmes.append(req.programme)
        years.append(req.courseyear)
        advisors.append(req.advisor)
        circumstances.append(req.circumstances)
        periods.append(req.dates_affected)
        latereasons.append(req.latereason)
        isDASS.append(req.DASS)
        evidences.append(req.evidence)
        toldsupervi.append(req.superinformed)
        supervisors.append(req.supervisor)
        onT4Visas.append(req.T4Visa)
        proposedDL.append(req.proposedDL)
        evidencesumm.append(req.evidencesummary)
        outcomes.append("Pending Outcome...")
        emailtypes.append("...")
        sendbyperson.append("Pending Outcome...")
        notes.append("...")
        panelnotes.append("...")
        outcomesent.append("Pending Send...")
        outcomedate.append("Pending Send...")

        # The assessment Names and Codes applied for are to be presented in the same cell, but
        # on separate lines. This is achieved here by joining the lists of assessments together
        # with newlines
        code_string:      str = "\n".join(req.asm_codes)
        name_string:      str = "\n".join(req.asm_names)
        other_string:     str = "\n".join(req.other_asm)
        isresub_string:   str = "\n".join(req.asm_is_resub)
        resubdate_string: str = "\n".join(req.asm_resubdate)
        status_string:    str = "\n".join(req.asm_resubstatus)
        
        asmcodes.append(code_string)
        asmnames.append(name_string)
        otherinfo.append(other_string)
        isresubs.append(isresub_string)
        resubdates.append(resubdate_string)
        substats.append(status_string)

    # Column dictionary as explained above
    columns: Dict[str, any] = {'Date Submitted': subdates, 'Full Name': names, 'University Email': emails, 'Student ID Number': IDs,
                               'Are you applying for a postgraduate dissertation or research project?': PGRs, 'No. Assessments/Exams': N_affected,
                               'Division': divisions, 'Programme': programmes, 'Year': years, 'Unit Code': asmcodes,
                               'Assessment name and submission date': asmnames, 'Is this a resubmission (including date)?': isresubs,
                               'Other Assessment Info.': otherinfo,
                               'Submission Status': substats, 'Academic Advisor(s)': advisors, 'Reason for Mitigation': circumstances, 'Period Affected': periods,
                               'Late Application - Reason': latereasons, 'DASS Registration': isDASS, 'Evidence Declaration': evidences,
                               'Supervisor Aware?': toldsupervi, 'Supervisor Name': supervisors, 'Tier 4 Visa': onT4Visas, 'Proposed New Deadline': proposedDL,
                               'Evidence Summary': evidencesumm, 'Outcome': outcomes, 'Email Type': emailtypes, 'To Be Sent By (Initials)...': sendbyperson,
                               'Notes': notes, 'Panel Notes': panelnotes, 'Outcome Sent to Student': outcomesent, 'Outcome Sent Date': outcomedate}

    # Create the dataframe from the column dictionary, including the number of unique students
    # making requests in the Excel spreadsheet sheet name
    dataframe: DataFrame = pandas.DataFrame(columns)
    sheetname: str = f"Mitigating Circumstances ({len(requests)})"

    # Using the .set_column() method of the ExcelWriter, it is possible to change the formatting and size of 
    # columns in the output spreadsheet.
    # All of the column widths need to be changed to some degree to tidy the spreadsheet and ensure that the
    # unique Assessment Names and Codes appear on separate lines correctly, as explained above - specifically,
    # these are the columns that will need to be resized manually, so the column names are listed here to
    # check against while writing to the spreadsheet.
    # manual_resize_fudge is simply how many additional pixels these columns should be widened by to ensure
    # that this isn't too tight
    manual_resize_fudge: int = 10
    cols_to_manually_resize: List[str] = ["Unit Code",
                                          "Assessment name and submission date",
                                          "Is this a resubmission (including date)?",
                                          "Submission Status"]
    
    print(dataframe)

    # The xlsxwriter library is frequently missing, for some reason - check to see if it is available
    # here, and download with Pip if it is not found (this will only happen the first time that the program runs)
    # Bit of a sledgehammer solution, but it gets the job done for now
    # TODO: Is there a less shit way of doing this?
    try:
        from xlsxwriter import Workbook
    except ModuleNotFoundError as mnfe:
        os.system("python -m pip install xlsxwriter")

    with pandas.ExcelWriter(output, engine = 'xlsxwriter') as xlwriter:
        # Write the dataframe to the output spreadsheet, and get an instance of
        # the WorkBook so that we can begin reformatting the spreadsheet as needed
        dataframe.to_excel(xlwriter, sheet_name = sheetname, index = False)
        workbook = xlwriter.book

        # To ensure that the newlines ("\n") function correctly, text wrapping needs to be
        # enabled for the manually-resized Assessment Name / Code / Status columns
        cellformat = workbook.add_format({'text_wrap': True})

        for colname in dataframe:
            column_index = dataframe.columns.get_loc(colname)

            # If this is one of those columns, then find the length of the longest line that the
            # cell will need to accomodate using required_column_width(), and add on the extra
            # fudge space - then use set_column() to resize this column and enable text wrapping
            if colname in cols_to_manually_resize:
                column_width = required_column_width(dataframe, colname) + manual_resize_fudge
                xlwriter.sheets[sheetname].set_column(column_index, column_index, column_width, cell_format = cellformat)

            # Otherwise we don't need to care about the within-cell newlines and can do this
            # automagically - simply find the longest string in any cell of this column and set
            # the width to that (or, if the *name* of the column is longer, simply use that
            # instead) - note again that these columns don't need to care about newlines within
            # cells and so I do not enable text wrapping for these
            else:
                column_width = max(dataframe[colname].astype(str).map(len).max(), len(colname))
                xlwriter.sheets[sheetname].set_column(column_index, column_index, column_width)

    print("    ...Done!")


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >


def run_startup_checks(qualtrics: str, output: str, logging: bool, verbose: bool) -> None:
    startup_error: bool = False
    info_message: str = date_today()

    # Firstly, check whether the user selected an output directory. If not, prompt them
    # again to select one. If, for whatever reason, they select 'No' here, then close down
    # the window and exit the program. Return 0 as this is not technically an error.
    if not output:
        result = tk.messagebox.askquestion("Select Directory", "No Output folder was selected!\nSelect one now?", icon = "warning")
        if result == "yes":
            select_output_folder_window()
            output = output_directory_entry.get()
        else:
            destroy_window()
            exit(0)

    # As with the Output directory above, we also prompt the user to select a Qualtrics file
    # if they have not done this already, and exit the program if they choose not to do so.
    if not qualtrics:
        result = tk.messagebox.askquestion("Select Spreadsheet", "No Qualtrics Spreadsheet was selected!\nSelect one now?", icon = "warning")
        if result == "yes":
            select_spreadsheet_file_window()
            qualtrics = input_requests_entry.get()
        else:
            destroy_window()
            exit(0)
    
    # Second, ensure that the input Qualtrics file is actually of the expected type (an Excel .xlsx file, or 
    # a Comma-Separated .csv file), and then ensure that the path actually points to a file that exists.
    # Show a helpful error window and exit the program if any of these checks fail.
    try:
        _, extension = os.path.splitext(qualtrics)

        if not extension in [".xlsx", ".csv"]:
            filetype_error_message: str = f"File '{os.path.basename(qualtrics)}' is not of the required type!\nCompatible files are '.xlsx' or '.csv'"
            tk.messagebox.showinfo(title = "Error!", message = filetype_error_message)
            startup_error = True

    except Exception as expt:
        print(f"Filepath error!\n  > Exception: '{expt}'")
        file_error_message: str = f"Path to file '{os.path.basename(qualtrics)}' is malformed!\nCheck the filepath and retry."
        tk.messagebox.showinfo(title = "Error!", message = file_error_message)
        startup_error = True

    if object_exists(qualtrics, suppress = True):
        info_message = info_message + f"\nQualtrics file '{os.path.basename(qualtrics)}': Found"
    else:
        info_message = info_message + f"\nQualtrics file '{os.path.basename(qualtrics)}' missing! Check path and retry..."
        startup_error = True

    # Finally, check if the selected Output directory currently exists. If it does
    # not, then display a message box asking the user if they would like it to be
    # created. If, for some reason, they select 'no', then this is considered a
    # startup error which will cause the program to be shut down
    if object_exists(output, suppress = True):
        info_message = info_message + f"\nOutput directory '{output}': Found"
    else:
        info_message = info_message + f"\nNew Output directory '{output}' was created."
        result = tk.messagebox.askquestion("Create Directory", "Create New Output Directory?", icon = "warning")
        if result == 'yes':
            print("New folder on the way!")
            os.mkdir(output)
        else:
            print("No new folder will be created.\nExiting...")
            startup_error = True
            
    # Once complete, check if we hit any major points of failure and exit if the program
    # now if so. Otherwise, display the accrued startup information message to the user.            
    if startup_error:
        destroy_window()
        exit(1)
    else:
        showinfo("Startup Information...", message = info_message)


# - - - - - - >


def main() -> None:
    # Run some basic startup checks and display results to the user in an information window
    # Ensure that the input Qualtrics file exists, that the output folder exists / can be
    # created, and whether the user has selected logging and / or verbose running.
    # Check if the user has requested logging of information as the program runs, and create
    # this log.txt file (if it does not already exist for some reason)
    run_startup_checks(input_requests_entry.get(), output_directory_entry.get(), write_logfile_flag.get(), display_running_information.get())
    logfile: str = create_log_if_requested(output_directory_entry.get(), write_logfile_flag.get())
    logging: bool = write_logfile_flag.get()
    logcount: Counter = Counter()
    _, extension = os.path.splitext(input_requests_entry.get())

    if logging:
        log_string(logfile, f"Starting up: [{current_datetime()}]", logcount)

    # The Uni-controlled laptops do not have Openpyxl installed, but Pandas seems to require
    # this in order to *read* files - so, we need to ensure this is available here and install
    # it using the command-line if it is missing
    try:
        import openpyxl
    except ModuleNotFoundError as mnfe:
        command: str = "python -m pip install openpyxl"
        if logging:
            log_string(logfile, f"openpyxl dependency missing. Installing using command '{command}'", logcount)
        os.system(command)

    # Read the raw data from the Qualtrics output into a DataFrame. 
    # Data can be read from either a .csv or a .xlsx file.
    # If the file is a .xlsx, then the sheet containing this should have a pre-specified name ("Sheet0").
    # If this is not present in the Excel file then assume it is contained in the first sheet
    try:
        if "csv" in extension:
            qualtrics: DataFrame = pandas.read_csv(input_requests_entry.get())
        else:
            qualtrics: DataFrame = pandas.read_excel(input_requests_entry.get(), sheet_name = "Sheet0")
    except ValueError as verr:
        if logging:
            log_string(logfile, f"{verr}: Sheet name 'Sheet0' not in spreadsheet - using sheet index = 0 instead.", logcount)
        qualtrics: DataFrame = pandas.read_excel(input_requests_entry.get(), sheet_name = 0)
    

    # The new version of the Qualtrics output appears to contain some Qualtrics-specific junk in
    # Excel row 3 - regardless of whether the user has asked for this to be cleaned, we need to
    # check for it and ensure it is removed otherwise it will produce mess in the output
    qualtrics = drop_row_by_string(qualtrics, "ImportId")


    # Parse the raw Qualtrics output data into a list of StudentRequest instances, a class which
    # contains all of the information on a given students' application (Name, ID, Year and Programme,
    # Assessments applied for, Unit Codes, Circumstances leading to their application, etc.)
    requests:  List[StudentRequest] = build_student_requests(qualtrics, display_running_information.get(), logging, logfile, logcount)


    # Write the parsed Student Requests out to a spreadsheet, formatted following the "Tracker"
    output_filename: str = create_output_filename(output_directory_entry.get(), len(requests))
    
    print(f"Emitting to: {os.path.basename(output_filename)}")
    
    requests_to_spreadsheet(requests, output_filename)

    if logging:
        log_string(logfile, f"Closing down: [{current_datetime()}]", logcount)
    

# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# - - - - - - - - - - - - - - ! BZZT, BZZT WARNING - GLOBAL SCOPE ! - - - - - - - - - - - - - - - >
#  Set up the main GUI interface with buttons and text-boxes for entering parameters, files, etc. >
#       Specify functions to be called on button-press, and entry-point for the main program      >
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >

def select_spreadsheet_file_window():
    valid_filetypes = (("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*"))
    filename = filedialog.askopenfilename(title = "Select a spreadsheet...", filetypes = valid_filetypes)

    # Only show the information window if the user actually selects a file here - if they Cancel out,
    # this would otherwise just show a blank window which looks busted
    if filename:
        showinfo(title = "Reading file...", message = filename)

    input_requests_entry.delete(0, tk.END)
    input_requests_entry.insert(0, filename)

# - - - - - - >

def select_output_folder_window():
    output_location = filedialog.askdirectory()
    output_directory_entry.delete(0, tk.END)
    output_directory_entry.insert(0, output_location)

# - - - - - - >

def check_verbose_running_flag():
    check_flag_message: str = None

    if display_running_information.get():
        check_flag_message = "Additional information will be shown in the command window as the program runs.\nCheck for any issues / errors!"
        tk.messagebox.showinfo(title = "Information...", message = check_flag_message)

    del(check_flag_message)

# - - - - - - >

def check_logfile_flag():
    check_logfile_flag_message: str = None

    if write_logfile_flag.get():
        check_logfile_flag_message = "Each step of the cleanup procedure will be logged to a text file.\nYou will find this in your chosen Output folder."
        tk.messagebox.showinfo(title = "Information...", message = check_logfile_flag_message)

    del(check_logfile_flag_message)

# - - - - - - >

def check_alternative_output_flag():
    check_output_flag_message: str = None

    if alternative_output_format.get():
        check_output_flag_message = "Output spreadsheet will use the alternative formatting ('Option 2')\nThis means each assessment will be written to its own unique row."
        tk.messagebox.showinfo(title = "Information...", message = check_output_flag_message)
    
    del(check_output_flag_message)

# - - - - - - >

def check_clean_input_flag():
    check_cleaning_flag_message: str = None

    if delete_junk_rows.get():
        check_cleaning_flag_message = "Junk rows in Qualtrics output will be removed prior to processing (typically Excel Row 3)\n"
        check_cleaning_flag_message = f"{check_cleaning_flag_message}If unsure, check the Qualtrics output to ensure this is right for your data."
        tk.messagebox.showinfo(title = "Information...", message = check_cleaning_flag_message)

    del(check_cleaning_flag_message)

# - - - - - - >

def show_help_windows():
    pass

# - - - - - - >

def destroy_window():
    print(f"[{current_datetime()}] Closing down...")
    parent.destroy()


# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >
# Set up the main window - ensure the window always displays on top (-topmost), and disable
# resizing in the X and Y directions
parent = tk.Tk()
parent.title("Mitigating Circumstances")
parent.geometry("360x375")
parent.call('wm', 'attributes', '.', '-topmost', '1')
parent.resizable(width = False, height = False)


# Attempt to apply a background image to the main window
# By default, there *should* be a folder named "Assets" in the same directory as this
# source file. If the file does exist in this subfolder, then place this on the parent
# window. Otherwise, do nothing / leave the parent window background plain grey
current_directory: str = os.getcwd()
previous_directory: str = os.path.dirname(current_directory)
background_imagepath: str = os.path.join(previous_directory, "Mitigations", "Assets", "background2.png")

if object_exists(background_imagepath, suppress = False):
    background_image = tk.PhotoImage(file = background_imagepath)
    background_label = tk.Label(parent, image = background_image)
    background_label.place(x = 0, y = 0)


# Entry box & button for the filepath to the mitigating circumstances Excel file
# this should be the raw file as emitted by Qualtrics
input_requests_label = tk.Label(parent, text = "MitCircs Qualtrics Input Excel File:")
input_requests_label.pack()
input_requests_entry = tk.Entry(parent)
input_requests_entry.pack()
input_requests_button = tk.Button(parent, text = "Select", command = select_spreadsheet_file_window)
input_requests_button.pack()


# Entry box & button for the path to the directory where the cleaned-up / processed
# output spreadsheet should be written to. If no Output location is selected, then
# this will default to the same folder that contains the Mitigating Circumstances file
output_directory_label = tk.Label(parent, text = "Output Folder Location:")
output_directory_label.pack()
output_directory_entry = tk.Entry(parent)
output_directory_entry.pack()
output_directory_button = tk.Button(parent, text = "Browse...", command = select_output_folder_window)
output_directory_button.pack()


# The "verbose" option - causes the program to display iteration-by-iteration information
# printed to the terminal as the request-builder and cleanup functions run. Only really
# relevant during debugging and while adding program features so can be enabled / disabled
# with this checkbox to run the program "quietly"
display_running_information = tk.BooleanVar()
display_running_information_checkbox = tk.Checkbutton(parent, text = "Display Information as Program Runs?",
                                                      variable = display_running_information, onvalue = True, offvalue = False,
                                                      command = check_verbose_running_flag)
display_running_information_checkbox.pack()


# Useful information on each request can be written to a logfile - the user can select
# here whether they would like this to be done. If so, a logfile will be created in the
# specified output folder with the current date and time as filename
write_logfile_flag = tk.BooleanVar()
write_logfile_flag_checkbox = tk.Checkbutton(parent, text = "Write Cleanup Info. to Logfile?",
                                             variable = write_logfile_flag, onvalue = True, offvalue = False,
                                             command = check_logfile_flag)
write_logfile_flag_checkbox.pack()


# Two possible formats for the output spreadsheet were provided - in the first, each
# assessment that the student applies for is written to a separate cell in the output
# sheet, with their name and identifying information on only the top row.
# In the second, all of the assessments that the student applies for are written into
# the *same* cell so that everything is kept on one row, with assessments separated
# only by newlines within the cell.
# Apparently the second format is preferable, but this toggle allows the user to select
# the first, separate-row output format if they would prefer
alternative_output_format = tk.BooleanVar()
alternative_output_format_checkbox = tk.Checkbutton(parent, text = "Use Alternative Output Format?",
                                                    variable = alternative_output_format, onvalue = True, offvalue = False,
                                                    command = check_alternative_output_flag)
#alternative_output_format_checkbox.pack()


# In the new (Nov. 2024) test data, there is an additional 3rd row below the header
# which appears to basically contain junk from the Qualtrics export. This option
# indicates whether we should attempt to search for and delete this junk row - if
# not, we can simply continue with the extraction as expected (1: Rowname, 2: Header,
# 3: Start of data rows).
# NOTE: As of the new Q4 2024 / Q1 2025 version this is automatically checked for and
#       deleted if found - there should be no reason for the user to think about this
#       themselves, so this code will be commented out until I'm absolutely sure it
#       is safe to be removed
delete_junk_rows = tk.BooleanVar()
delete_junk_rows_checkbox = tk.Checkbutton(parent, text = "Clean Qualtrics junk from input?",
                                           variable = delete_junk_rows, onvalue = True, offvalue = False,
                                           command = check_clean_input_flag)
# delete_junk_rows_checkbox.pack()


# Show some useful help windows and images to show expected file and formatting.
# If the program is reporting issues and errors with the inputs, then start here!
help_window_button = tk.Button(parent, text = "Show Help", command = show_help_windows)
#help_window_button.pack()


# Run the MitCircs processing program with the inputs specified
# Firstly checks to ensure that all required inputs are present and can be found
# on the system. If any inputs are missing or not provided, message-boxes will
# inform you of the problem and then return to the parent main-loop for you to fix them
run_main_button = tk.Button(parent, text = "Run!", command = main)
run_main_button.pack()


# Destroy parent window and children, exiting the program
quit_button = tk.Button(parent, text = "Exit...", command = destroy_window)
quit_button.pack()


# Begin running the parent main-loop, awaiting inputs
parent.mainloop()