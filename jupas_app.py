import streamlit as st
import pandas as pd
import os
import json
import re
import math # Import math
import traceback

# --- Constants ---
EXCEL_FILE = "jupas_programs.xlsx"
DEFAULT_CHOICES_CSV = "student_choices.csv"
PROGRAM_ID_COL_CLEANED = "program_id"
WEIGHTING_COL_CLEANED = "subject_weightings"
REGNO_COL = "regno"
FIXED_CLASS_OPTIONS = ["", "6A", "6B", "6C", "6D"]
CORE_ENG = "english language"
CORE_MATH = "mathematics compulsory part"
CORE_CHI = "chinese language"
CORE_CS = ["citizenship and social development", "liberal studies"]
SCIENCE_ELECTIVES = ["biology", "chemistry", "physics", "mathematics m1", "mathematics m2"]
ENG_ELECTIVES = ["biology", "chemistry", "physics", "mathematics m1", "mathematics m2", "ict"] #for HKUST Engineering


# --- Helper Functions --- (Defined FIRST)

import json
import pandas as pd # Assuming pandas is used for pd.isna

import json
import pandas as pd
import re # Import re again, might be useful

def parse_weights_string(weight_str):
    """
    Cleans and parses a weight string into a Python dictionary.
    Handles various quote issues and formatting errors found in data.
    Converts dictionary keys to lowercase.
    """
    if not isinstance(weight_str, str) or pd.isna(weight_str):
        return {}

    original_input = weight_str
    text = weight_str.strip()

    # --- Log initial state ---
    # print(f"DEBUG Parse Start: Input='{original_input}', Stripped='{text}'")

    # 1. Strip potential TRIPLE quotes
    if len(text) >= 6 and text.startswith('"""') and text.endswith('"""'):
        text = text[3:-3].strip()

    # 2. Strip potential surrounding DOUBLE quotes IF content looks like dict {}
    if len(text) >= 2 and text.startswith('"') and text.endswith('"'):
        content_inside = text[1:-1]
        if content_inside.strip().startswith('{') and content_inside.strip().endswith('}'):
            text = content_inside

    # --- Pre-JSON Fixes ---
    # 2.5: Replace internal "" with " (Handles escaped quotes)
    if '""' in text:
        text = text.replace('""', '"')
        # print(f"DEBUG Parse Step 2.5: Replaced internal quotes -> '{text}'")

    # 2.6: Fix missing quote before colon (e.g., "economics: -> "economics":)
    # Use regex to find keys missing the quote just before the colon
    text = re.sub(r'(?<=\{|,)\s*([a-zA-Z0-9_]+)\s*:', r'"\1":', text)
    # print(f"DEBUG Parse Step 2.6: Fixed missing quotes -> '{text}'")

    # 2.7: Fix period used as delimiter instead of comma (e.g., "key": val. "key2")
    # Look for quote-colon-space-number-period-space-quote pattern
    text = re.sub(r'(": \d+)(\.)(\s*")', r'\1,\3', text) # Handles integer followed by . "
    text = re.sub(r'(": \d+\.\d+)(\.)(\s*")', r'\1,\3', text) # Handles float followed by . "
    # print(f"DEBUG Parse Step 2.7: Fixed period delimiter -> '{text}'")

    # 2.8: Fix comma used as decimal separator (e.g., 1,5 -> 1.5)
    # Only replace comma if it's between two digits
    text = re.sub(r'(\d),(\d)', r'\1.\2', text)
    # print(f"DEBUG Parse Step 2.8: Fixed comma decimal -> '{text}'")
    # --- End Pre-JSON Fixes ---


    if text == '{}': return {}
    # Check structure again after fixes
    if not text.strip().startswith('{') or not text.strip().endswith('}'):
         # print(f"Warning: String does not look like JSON dict after cleaning. Input='{original_input}', Final Text='{text}'")
         return {}

    # --- Log final string before parsing ---
    # print(f"DEBUG Parse Final Attempt: Parsing '{text}'")

    # 3. Attempt to parse
    try:
        weights_dict = json.loads(text)
        if not isinstance(weights_dict, dict):
             print(f"Warning: Parsed result is not a dict! Input='{original_input}', Parsed='{text}' -> {weights_dict}")
             return {}
        # print(f"DEBUG Parse Success: Input='{original_input}' -> Parsed='{text}' -> Result={weights_dict}")
        return {str(k).lower(): v for k, v in weights_dict.items()}

    except json.JSONDecodeError as e:
        print(f"ERROR: json.loads failed for Input='{original_input}'. Attempted to parse: '{text}'. Error: {e}")
        return {}
    except Exception as e_gen:
         print(f"ERROR: Unexpected error during parsing. Input='{original_input}'. Attempted text: '{text}'. Error: {e_gen}")
         return {}

# --- Define DSE Grade to Points Mappings ---

POINTS_DEFAULT = {"5**": 7.0, "5*": 6.0, "5": 5.0, "4": 4.0, "3": 3.0, "2": 2.0, "1": 1.0, "u": 0.0, "a": 0.0, "": 0.0, None: 0.0} # HKBU, LingU, EdUHK, Others
POINTS_SCALE_8_5 = {"5**": 8.5, "5*": 7.0, "5": 5.5, "4": 4.0, "3": 3.0, "2": 2.0, "1": 1.0, "u": 0.0, "a": 0.0, "": 0.0, None: 0.0} # CityU, CUHK, HKU, PolyU
POINTS_HKUST = {"5**": 8.5, "5*": 7.0, "5": 5.5, "4": 4.0, "3": 3.0, "2": 2.0, "1": 0.0, "u": 0.0, "a": 0.0, "": 0.0, None: 0.0} # HKUST (L1=0)

# --- Helper Function 2: Select Point Map Based on Institution ---
def get_point_map_for_institution(institution_name_str):
    """
    Selects the correct DSE points map dictionary based on institution name.

    Args:
        institution_name_str (str): The name of the institution.

    Returns:
        dict: The dictionary mapping grades to points for that institution.
    """
    # Add empty string/None defaults to all maps for safety if not already done
    for map_dict in [POINTS_DEFAULT, POINTS_SCALE_8_5, POINTS_HKUST]:
        map_dict[""] = 0.0
        map_dict[None] = 0.0 # Handle potential None type if data has it

    if not institution_name_str or not isinstance(institution_name_str, str):
        # print("DEBUG: Institution name missing or invalid, using DEFAULT points map.") # Optional debug
        return POINTS_DEFAULT # Default if name is missing or not a string

    inst_lower = institution_name_str.lower().strip()

    # Use specific checks for known variations
    if "cityu" == inst_lower or "city university" in inst_lower:
        # print(f"DEBUG: Using SCALE_8_5 for {institution_name_str}")
        return POINTS_SCALE_8_5
    elif "hkbu" == inst_lower or "baptist university" in inst_lower:
        # print(f"DEBUG: Using DEFAULT for {institution_name_str}")
        return POINTS_DEFAULT
    elif "polyu" == inst_lower or "polytechnic university" in inst_lower:
        # print(f"DEBUG: Using SCALE_8_5 for {institution_name_str}")
        return POINTS_SCALE_8_5
    elif "cuhk" == inst_lower or "chinese university" in inst_lower:
        # print(f"DEBUG: Using SCALE_8_5 for {institution_name_str}")
        return POINTS_SCALE_8_5
    elif "hku" == inst_lower or "university of hong kong" == inst_lower:
        # print(f"DEBUG: Using SCALE_8_5 for {institution_name_str}")
        return POINTS_SCALE_8_5
    elif "hkust" == inst_lower or "university of science and technology" == inst_lower or "ust" == inst_lower:
        # print(f"DEBUG: Using HKUST map for {institution_name_str}")
        return POINTS_HKUST
    elif "lingu" == inst_lower or "lingnan university" in inst_lower:
        # print(f"DEBUG: Using DEFAULT for {institution_name_str}")
        return POINTS_DEFAULT
    elif "eduhk" == inst_lower or "education university" in inst_lower:
        # print(f"DEBUG: Using DEFAULT for {institution_name_str}")
        return POINTS_DEFAULT
    else:
        # Assuming any other institution uses the default scale
        # print(f"DEBUG: Institution '{institution_name_str}' not matched, using DEFAULT points map.") # Optional debug
        return POINTS_DEFAULT

# --- Helper Function 3: Dynamic Grade to Points ---
def grade_to_points_dynamic(single_grade_string, point_map):
    """
    Converts a single DSE grade string into points using a provided map.
    Returns 0.0 for unrecognized grades within that map or for CS='A'.

    Args:
        single_grade_string (str): The DSE grade string (e.g., "5*", "A").
        point_map (dict): The specific points map dictionary to use.

    Returns:
        float: The corresponding points (as a float).
    """
    if not isinstance(single_grade_string, str):
        # print(f"DEBUG grade_to_points: Input not string '{single_grade_string}', returning 0.0") # Optional debug
        return 0.0

    # Standardize: lowercase and remove whitespace (keys in maps are lowercase)
    standardized_grade = single_grade_string.strip().lower()

    # Handle specific cases before dictionary lookup
    if not standardized_grade:
        # print(f"DEBUG grade_to_points: Empty standardized grade from '{single_grade_string}', returning 0.0") # Optional debug
        return 0.0
    if standardized_grade == "a":
        # print(f"DEBUG grade_to_points: Grade 'a' (Attained), returning 0.0") # Optional debug
        return 0.0 # Attained = 0 points for calculation

    # Lookup using the provided map, default to 0.0 if not found
    point_value = point_map.get(standardized_grade, 0.0)

    # Ensure float return type
    # print(f"DEBUG grade_to_points: Grade '{standardized_grade}' -> Points {float(point_value)}") # Optional debug
    return float(point_value)

# --- Subject Name Mapping and Weight Key Function ---
SUBJECT_NAME_MAP_TO_WEIGHT_KEY = {
    # Core
    'chinese language': 'chinese', 'chineselang': 'chinese',
    'english language': 'english', 'eng': 'english',
    'mathematics compulsory part': 'math', 'math compulsory': 'math',
    'citizenship and social development': 'citizenship', 'cs': 'citizenship', 'liberal studies': 'citizenship',

    # Extended Math
    'mathematics m1': 'm1',
    'mathematics m2': 'm2',

    # Common Electives (add more as needed based on weighting keys found)
    'physics': 'physics', 'phy': 'physics',
    'chemistry': 'chemistry', 'chem': 'chemistry',
    'biology': 'biology', 'bio': 'biology',
    'history': 'history', 'hist': 'history',
    'chinese history': 'chinese_history', 'chist': 'chinese_history',
    'economics': 'economics', 'econ': 'economics',
    'geography': 'geography', 'geog': 'geography',
    'bafs': 'bafs',
    'ict': 'ict',
    'visual arts': 'visual arts', 'va': 'visual arts', 'visual_arts': 'visual arts',
    'music': 'music',
    'physical education': 'physical_education', 'pe': 'physical_education',
    'chinese literature': 'chinese_literature', 'chinlit': 'chinese_literature',
    'literature in english': 'literature_in_english', 'englit': 'literature_in_english',
    'tourism and hospitality studies': 'tourism_hospitality_studies', 'thm': 'tourism_hospitality_studies',
    'design and applied technology': 'design_applied_technology', 'design_tech': 'design_applied_technology',
    'combined science': 'combined_science',
    'integrated science': 'integrated_science'
    # Add other electives...
}
def get_weighting_key(subject_name_full):
    """Finds the potential key for weight lookup from the full subject name."""
    name_lower = subject_name_full.lower().strip()
    return SUBJECT_NAME_MAP_TO_WEIGHT_KEY.get(name_lower, name_lower) # Default to lowercase name if no specific mapping


# --- Data Loading/Preparation Functions --- (Defined AFTER helpers)

@st.cache_data # Caches the result
def load_and_prepare_jupas_data(filepath=EXCEL_FILE):
    """
    Loads JUPAS data from Excel and performs all cleaning steps.
    Returns the prepared DataFrame or None if loading fails.
    Includes print statements for initial loading confirmation (will show in terminal).
    """
    print(f"Attempting to load and prepare JUPAS data from: {filepath}") # Shows in terminal
    if not os.path.exists(filepath):
        st.error(f"Fatal Error: JUPAS Program file not found at '{filepath}'. Application cannot start.")
        print(f"Fatal Error: JUPAS Program file not found at '{filepath}'.")
        return None
    try:
        df = pd.read_excel(filepath)
        print(f"Successfully loaded {len(df)} rows and {len(df.columns)} columns from Excel.")
    except Exception as e:
        st.error(f"Fatal Error: Could not load JUPAS Program file '{filepath}'. Reason: {e}")
        print(f"Fatal Error: Could not load JUPAS Program file '{filepath}'. Reason: {e}")
        return None
    
    # --- Cleaning Steps ---
    print("Starting data cleaning...")

    # --- Add Debugging Here ---
    TEST_ID_FOR_WEIGHTS = "JS6080" # Use the ID you are testing
    print(f"\nDEBUG: Checking original weighting string for {TEST_ID_FOR_WEIGHTS}:")
    try:
        # Use the ORIGINAL Program ID column name BEFORE cleaning
        # And the ORIGINAL Weighting column name
        original_program_id_col = "Program ID" #<-- Use original name from Excel
        original_weighting_col = "Subject Weightings" #<-- Use original name from Excel

        # Find the row using the original column name and test ID (ensure case matches Excel or convert test ID)
        # Assume Test ID is already uppercase from constant definition
        original_weight_str = df.loc[df[original_program_id_col].astype(str).str.upper() == TEST_ID_FOR_WEIGHTS, original_weighting_col].iloc[0]
        print(f"  Original String: '{original_weight_str}' (Type: {type(original_weight_str)})")
    except IndexError:
         print(f"  Error: Program ID {TEST_ID_FOR_WEIGHTS} not found using original column name '{original_program_id_col}'.")
    except KeyError as e:
        print(f"  Error: Column '{e}' not found when searching original data.") # Will show which column is missing
    except Exception as e:
        print(f"  Could not find original string for {TEST_ID_FOR_WEIGHTS}: {e}")
    
    
    # Clean Column Names
    df.columns = df.columns.str.replace(' ', '_', regex=False)\
                           .str.replace('[^A-Za-z0-9_]+', '', regex=True)\
                           .str.lower()
    print("Column names cleaned.")

    # Parse Subject Weightings
    if WEIGHTING_COL_CLEANED in df.columns:
        print(f"Parsing '{WEIGHTING_COL_CLEANED}'...")
        # This call works because parse_weights_string is defined above
        df['subject_weightings_dict'] = df[WEIGHTING_COL_CLEANED].fillna('').apply(parse_weights_string)
        print("Weightings parsed.")
    else:
        print(f"Warning: Weighting column '{WEIGHTING_COL_CLEANED}' not found.")

    # --- DEBUG BLOCK ---
        print(f"\nDEBUG: Checking parsed weighting dict for {TEST_ID_FOR_WEIGHTS}:")
        try:
            # Find the parsed dict using the CLEANED column name
            # Need to find it before setting the index
            parsed_dict = df.loc[df[PROGRAM_ID_COL_CLEANED].astype(str).str.upper() == TEST_ID_FOR_WEIGHTS, 'subject_weightings_dict'].iloc[0]
            print(f"  Parsed Dict: {parsed_dict} (Type: {type(parsed_dict)})")
            if not parsed_dict:
                 print(f"  >>> PARSED DICT IS EMPTY - PARSING LIKELY FAILED for original string above.")
        except IndexError:
             print(f"  Error: Program ID {TEST_ID_FOR_WEIGHTS} not found using cleaned column name '{PROGRAM_ID_COL_CLEANED}' before setting index.")
        except KeyError as e:
             print(f"  Error: Column '{e}' not found when searching cleaned data before index.")
        except Exception as e:
             print(f"  Could not find parsed dict for {TEST_ID_FOR_WEIGHTS}: {e}")
    # --- END SECOND DEBUG BLOCK ---

    # Set Program ID as Index
    if PROGRAM_ID_COL_CLEANED in df.columns:
        print(f"Setting index to '{PROGRAM_ID_COL_CLEANED}'...")
        if df[PROGRAM_ID_COL_CLEANED].isna().any():
             print(f"Warning: Missing values found in '{PROGRAM_ID_COL_CLEANED}'.") # Terminal warning
        if df[PROGRAM_ID_COL_CLEANED].duplicated().any():
             print(f"Warning: Duplicate values found in '{PROGRAM_ID_COL_CLEANED}'.") # Terminal warning
        df[PROGRAM_ID_COL_CLEANED] = df[PROGRAM_ID_COL_CLEANED].astype(str).str.upper()
        # Use errors='ignore' with set_index if duplicates are expected/okay, otherwise it errors
        try:
            df.set_index(PROGRAM_ID_COL_CLEANED, inplace=True, verify_integrity=False) # verify_integrity=False allows duplicates
            print("Index set.")
        except Exception as e:
             st.error(f"Fatal Error setting index: {e}")
             print(f"Fatal Error setting index: {e}")
             return None
    else:
         st.error(f"Fatal Error: Cannot find Program ID column '{PROGRAM_ID_COL_CLEANED}'.")
         print(f"Fatal Error: Cannot find Program ID column '{PROGRAM_ID_COL_CLEANED}'.")
         return None

    # Convert Numeric Columns
    print("Converting numeric columns...")
    cols_to_convert = [
        'intake_2025', 'band_a_applicants_2024', 'competition_ratio', 'admissions_2024',
        'uq_chinese', 'uq_english', 'uq_math', 'uq_citizenship', 'uq_m1m2',
        'uq_elective_1', 'uq_elective_2', 'uq_elective_3', 'uq_elective_4', 'uq_total',
        'm_chinese', 'm_english', 'm_math', 'm_citizenship', 'm_m1m2',
        'm_elective_1', 'm_elective_2', 'm_elective_3', 'm_elective_4', 'm_total',
        'lq_chinese', 'lq_english', 'lq_math', 'lq_citizenship', 'lq_m1m2',
        'lq_elective_1', 'lq_elective_2', 'lq_elective_3', 'lq_elective_4', 'lq_total'
    ]
    converted_count = 0
    for col in cols_to_convert:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            converted_count += 1
    print(f"{converted_count} columns converted to numeric.")

    print("JUPAS Data preparation complete.") # Shows in terminal
    return df


def load_student_choices_from_upload(uploaded_file):
    """
    Loads student JUPAS choices from an uploaded file object.
    """
    if uploaded_file is None:
        return None
    try:
        choices_df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
        # Clean column names
        choices_df.columns = choices_df.columns.str.replace(' ', '_', regex=False)\
                                       .str.replace('[^A-Za-z0-9_]+', '', regex=True)\
                                       .str.lower()
        # Ensure RegNo column exists
        global REGNO_COL # Ensure we use the global constant name
        # Try both 'regno' and 'registration_no' or similar common variations?
        possible_regno_cols = ['regno', 'reg_no', 'registration_no', 'student_id']
        found_regno_col = None
        for col in possible_regno_cols:
            if col in choices_df.columns:
                found_regno_col = col
                REGNO_COL = col # Update global constant if found differently
                print(f"Found registration number column as: {REGNO_COL}")
                break

        if not found_regno_col:
             st.error(f"Error in uploaded file: Could not find a registration number column (tried {possible_regno_cols}).")
             return None

        print(f"Successfully loaded {len(choices_df)} student records from uploaded file.") # Terminal log
        return choices_df
    except Exception as e:
        st.error(f"Error loading student choices CSV: {e}")
        print(f"Error loading student choices CSV: {e}") # Terminal log
        return None

# --- Reusable Helper Function for Sorting ---
def get_sorted_weighted_scores(weighted_dict, exclude_cs=True, exclude_subjects=None):
    """
    Gets weighted scores from a dictionary, optionally excludes CS and other subjects,
    and returns them sorted descending. Handles non-numeric scores.
    """
    scores_list = []
    if exclude_subjects is None:
        exclude_subjects = []

    # Ensure CS is always in the exclude list if flag is True
    cs_names = ["citizenship and social development", "liberal studies"] # Lowercase
    if exclude_cs:
        exclude_subjects.extend(cs_names)
    # Normalize exclude list to lowercase
    exclude_subjects_lower = {name.lower().strip() for name in exclude_subjects}

    for subject_name, score in weighted_dict.items():
        name_lower = subject_name.lower().strip()
        # Check exclusion list
        if name_lower in exclude_subjects_lower:
            continue
        # Ensure score is numeric before adding
        if isinstance(score, (int, float)) and pd.notna(score):
            scores_list.append(score)
        # else: ignore non-numeric scores

    return sorted(scores_list, reverse=True)

# --- Other Main Functions ---

# Modify display_program_details to use Streamlit elements
def display_program_details(dataframe, user_program_id):
    """
    Displays key details for a specific program ID using Streamlit elements.
    """
    if dataframe is None:
        st.error("Error: JUPAS DataFrame not loaded or prepared.")
        return

    standardized_id = str(user_program_id).upper().strip()

    if standardized_id in dataframe.index:
        try:
            program_data = dataframe.loc[standardized_id]

            st.subheader(f"Program Details: {standardized_id}")

            def get_detail_str(data_series, col_name, format_spec=None, na_value="N/A"):
                value = data_series.get(col_name, pd.NA) # Safe get
                if pd.isna(value): return na_value
                if format_spec:
                    try: return format(value, format_spec)
                    except (TypeError, ValueError): return str(value)
                # Special handling for dictionary display?
                if isinstance(value, dict): return json.dumps(value) # Display dict as JSON string
                return str(value)

            # Use columns for better layout
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**Program Name:** {get_detail_str(program_data, 'program_name')}")
                st.markdown(f"**Institution:** {get_detail_str(program_data, 'institution')}")
                st.markdown(f"**Intake 2025:** {get_detail_str(program_data, 'intake_2025', format_spec='.0f')}")
                st.markdown(f"**Admissions (2024):** {get_detail_str(program_data, 'admissions_2024', format_spec='.0f')}")

            with col2:
                 st.markdown(f"**LQ Total (2024):** {get_detail_str(program_data, 'lq_total', format_spec='.2f')}")
                 st.markdown(f"**Median Total (2024):** {get_detail_str(program_data, 'm_total', format_spec='.2f')}")
                 st.markdown(f"**UQ Total (2024):** {get_detail_str(program_data, 'uq_total', format_spec='.2f')}")
                 st.markdown(f"**Competition Ratio (2024):** {get_detail_str(program_data, 'competition_ratio', format_spec='.2f')}")

            st.markdown(f"**Scoring Method:**")
            st.info(get_detail_str(program_data, 'scoring_method'))
            st.markdown(f"**Subject Weightings:**")
            weights_dict = program_data.get('subject_weightings_dict', {})
            if weights_dict:
                # Format dictionary items as "key: value" pairs and join them
                weights_str = ", ".join([f"{k}: {v}" for k, v in weights_dict.items()])
                st.info(weights_str) # Use st.text for a simple, single line display
            else:
                st.info("None specified")

        except Exception as e:
             st.error(f"An unexpected error occurred displaying details for '{standardized_id}': {e}")
    else:
        st.error(f"Error: Program ID '{standardized_id}' not found in the data.")


# --- Calculate Score function (University-Aware) ---
def calculate_admission_score(dataframe, target_program_id, student_results):
    """
    Calculates the estimated JUPAS admission score for a specific program
    based on the student's results, program rules, and UNIVERSITY-SPECIFIC
    DSE point scales. Includes common rules for HKU, HKUST, CUHK, BestN, 3C+2X.

    Args:
        dataframe (pd.DataFrame): Prepared JUPAS data.
        target_program_id (str): The program ID to calculate for.
        student_results (dict): Dict of {'Subject Name': 'Grade String', ...}.

    Returns:
        tuple: (float or None, str)
               - The calculated score (float), or None if calculation fails.
               - Details message (str) about calculation or error.
    """
    # --- Initial Checks and Data Retrieval ---
    if dataframe is None: return None, "Error: JUPAS DataFrame not loaded."
    if not isinstance(student_results, dict): return None, "Error: student_results must be a dictionary."
    standardized_id = str(target_program_id).upper().strip()
    if standardized_id not in dataframe.index: return None, f"Error: Program ID '{standardized_id}' not found."

    try:
        program_data_row = dataframe.loc[standardized_id]
        scoring_method_str = program_data_row.get('scoring_method', '')
        weightings_dict = program_data_row.get('subject_weightings_dict', {})
        if not isinstance(weightings_dict, dict): weightings_dict = {} # Ensure dict

        # --- GET INSTITUTION NAME ---
        institution_name = program_data_row.get('institution', '') # Get institution
        # --- END GET ---

        if pd.isna(scoring_method_str) or not scoring_method_str or "not available" in scoring_method_str.lower():
            return None, f"Error: Scoring method for '{standardized_id}' is not available."
        scoring_method_clean = scoring_method_str.lower().strip()

    except Exception as e: return None, f"Error retrieving program data for '{standardized_id}': {e}"

    # --- Step 6: Convert Grades to Points (using the CORRECT scale) ---
    # --- SELECT POINT MAP ---
    points_map_to_use = get_point_map_for_institution(institution_name)
    # --- END SELECT ---

    subject_points_dict = {}
    for subject_name, grade_string in student_results.items():
        # --- CALL DYNAMIC CONVERTER ---
        points = grade_to_points_dynamic(grade_string, points_map_to_use)
        # --- END CALL ---
        subject_points_dict[subject_name.strip()] = points # Use stripped name as key

    # --- Step 7a: Apply Weights ---
    weighted_scores_dict = {}
    calc_details_list = []
    for subject_name, points in subject_points_dict.items():
        subj_name_strip = subject_name.strip() # Already stripped above, but safe
        weighting_key = get_weighting_key(subj_name_strip)
        weight = weightings_dict.get(weighting_key, 1.0)
        if not isinstance(weight, (int, float)): weight = 1.0 # Ensure numeric weight
        weighted_score = points * weight
        weighted_scores_dict[subj_name_strip] = weighted_score # Use stripped name
        # Store details (show points with 1 decimal for clarity, esp. with 8.5 scales)
        detail_str = f"{subj_name_strip}[{points:.1f}"
        if weight != 1.0: detail_str += f"x{weight:.2f}"
        detail_str += f"={weighted_score:.2f}]"
        calc_details_list.append(detail_str)

    # Add institution scale info to details message
    initial_details_msg = f"Method: {scoring_method_str} (Using {institution_name or 'Default'} scale). Inputs: {', '.join(calc_details_list)}."

    # --- Step 7b: Selection Logic ---
    final_score = None
    details_msg = initial_details_msg # Start with input details

    try:
        # Use lowercase keys for consistent lookup within this logic block
        weighted_scores_dict_lower_keys = {k.lower(): v for k, v in weighted_scores_dict.items()}

        # --- Get Sorted OVERALL Scores (excluding CS by default) ---
        sorted_scores_all = get_sorted_weighted_scores(weighted_scores_dict_lower_keys, exclude_cs=True)
        num_scores_avail = len(sorted_scores_all)

        # --- Rule Implementations (Order Matters - Specific Before General) ---

        # 1. HKU BBA/Econ Style (Eng*1.5 + Math*1.5 + Best 3 Others + Bonus)
        is_hku_bba_6th = "+ 0.2 x 6th best subject" in scoring_method_clean
        is_hku_bba_7th = "+ 0.2 x 7th best subject" in scoring_method_clean
        has_hku_weights = "1.5 x eng" in scoring_method_clean and "1.5 x math" in scoring_method_clean

        if (is_hku_bba_6th or is_hku_bba_7th) and has_hku_weights:
            score_eng = weighted_scores_dict_lower_keys.get(CORE_ENG, 0)
            score_math = weighted_scores_dict_lower_keys.get(CORE_MATH, 0)
            other_subjects_scores = get_sorted_weighted_scores(
                weighted_scores_dict_lower_keys, exclude_cs=True, exclude_subjects=[CORE_ENG, CORE_MATH]
            )
            num_other_scores = len(other_subjects_scores)
            if num_other_scores < 3: return None, details_msg + " Error: Less than 3 required 'other' subjects available."

            sum_best_3_others = sum(other_subjects_scores[:3])
            bonus_multiplier = 0.2
            bonus_score = 0.0
            next_best_index = -1
            if is_hku_bba_6th and num_scores_avail >= 6: next_best_index = 5
            elif is_hku_bba_7th and num_scores_avail >= 7: next_best_index = 6

            if next_best_index != -1:
                 sixth_or_seventh_best_overall_score = sorted_scores_all[next_best_index]
                 bonus_score = sixth_or_seventh_best_overall_score * bonus_multiplier

            final_score = score_eng + score_math + sum_best_3_others + bonus_score
            details_msg += f" Selected: Eng({score_eng:.2f}) + Math({score_math:.2f}) + Sum(Best 3 Others)({sum_best_3_others:.2f})"
            if bonus_score != 0: details_msg += f" + Bonus({bonus_score:.2f})"
            details_msg += f" -> Score: {final_score:.2f}"
            return final_score, details_msg


        # 2. HKUST Science Style (Eng*1.5 + Math*1 + Best Sci Elec + 2 Best Other)
        elif ("eng x 1.5" in scoring_method_clean and
              "best science elective" in scoring_method_clean and
              "2 best other" in scoring_method_clean):
            score_eng = weighted_scores_dict_lower_keys.get(CORE_ENG, 0)
            score_math = weighted_scores_dict_lower_keys.get(CORE_MATH, 0)
            best_sci_score, best_sci_name = 0, None
            temp_best_score = -1
            for subj_name_lower in SCIENCE_ELECTIVES:
                current_score = weighted_scores_dict_lower_keys.get(subj_name_lower, -1)
                if current_score > temp_best_score:
                     temp_best_score = current_score
                     best_sci_score = current_score
                     best_sci_name = subj_name_lower
            if temp_best_score == -1: best_sci_score = 0

            exclude_list_sci = [CORE_ENG, CORE_MATH, best_sci_name] if best_sci_name else [CORE_ENG, CORE_MATH]
            other_scores_sci = get_sorted_weighted_scores(
                weighted_scores_dict_lower_keys, exclude_cs=True, exclude_subjects=exclude_list_sci
            )
            sum_best_2_others = sum(other_scores_sci[:2])
            final_score = score_eng + score_math + best_sci_score + sum_best_2_others
            details_msg += f" Selected: Eng({score_eng:.2f}) + Math({score_math:.2f}) + BestSci({best_sci_name or 'N/A'}={best_sci_score:.2f}) + Sum(Best 2 Others)({sum_best_2_others:.2f}) -> Score: {final_score:.2f}"
            return final_score, details_msg

        # 3. HKUST Engineering Style (Eng*2 + Math*2 + Best Specific Elec + 2 Best Other)
        elif ("eng x 2" in scoring_method_clean and "math x 2" in scoring_method_clean and
             ("best elective" in scoring_method_clean or "best elec" in scoring_method_clean) and
              "2 best other" in scoring_method_clean):
            score_eng = weighted_scores_dict_lower_keys.get(CORE_ENG, 0)
            score_math = weighted_scores_dict_lower_keys.get(CORE_MATH, 0)
            best_eng_elec_score, best_eng_elec_name = 0, None
            temp_best_eng_score = -1
            for subj_name_lower in ENG_ELECTIVES:
                current_score = weighted_scores_dict_lower_keys.get(subj_name_lower, -1)
                if current_score > temp_best_eng_score:
                     temp_best_eng_score = current_score
                     best_eng_elec_score = current_score
                     best_eng_elec_name = subj_name_lower
            if temp_best_eng_score == -1: best_eng_elec_score = 0

            exclude_list_eng = [CORE_ENG, CORE_MATH, best_eng_elec_name] if best_eng_elec_name else [CORE_ENG, CORE_MATH]
            other_scores_eng = get_sorted_weighted_scores(
                weighted_scores_dict_lower_keys, exclude_cs=True, exclude_subjects=exclude_list_eng
            )
            sum_best_2_others_eng = sum(other_scores_eng[:2])
            final_score = score_eng + score_math + best_eng_elec_score + sum_best_2_others_eng
            details_msg += f" Selected: Eng({score_eng:.2f}) + Math({score_math:.2f}) + BestEngElec({best_eng_elec_name or 'N/A'}={best_eng_elec_score:.2f}) + Best 2 Others({sum_best_2_others_eng:.2f}) -> Score: {final_score:.2f}"
            return final_score, details_msg

        # 4. HKUST Business Style (Eng*2 + Math*2 + Best 3 Other)
        elif "eng x 2" in scoring_method_clean and "math x 2" in scoring_method_clean and "best 3 other" in scoring_method_clean:
            score_eng = weighted_scores_dict_lower_keys.get(CORE_ENG, 0)
            score_math = weighted_scores_dict_lower_keys.get(CORE_MATH, 0)
            other_scores_bus = get_sorted_weighted_scores(
                 weighted_scores_dict_lower_keys, exclude_cs=True, exclude_subjects=[CORE_ENG, CORE_MATH]
            )
            sum_best_3_others = sum(other_scores_bus[:3])
            final_score = score_eng + score_math + sum_best_3_others
            details_msg += f" Selected: Eng({score_eng:.2f}) + Math({score_math:.2f}) + Sum(Best 3 Others)({sum_best_3_others:.2f}) -> Score: {final_score:.2f}"
            return final_score, details_msg

        # 5. CUHK Science Style (Best 5 from Eng*1.5, Math*1.5, BestSci*2, Others*1)
        elif "best sci elec x 2" in scoring_method_clean and "eng x 1.5" in scoring_method_clean and "math x 1.5" in scoring_method_clean:
            final_weighted_scores_cuhk = {}
            best_sci_name_cuhk, best_sci_base_points = None, -1
            # Find best sci based on BASE points (before weighting)
            for subj_name_lower in SCIENCE_ELECTIVES:
                 # Need original TitleCase name to look up in subject_points_dict
                 subj_title_case = subj_name_lower.title() # Basic title case conversion
                 # Handle specific cases like BAFS, ICT etc. if needed
                 if subj_name_lower == 'bafs': subj_title_case = 'BAFS'
                 if subj_name_lower == 'ict': subj_title_case = 'ICT'
                 # ... add others ...
                 base_points = subject_points_dict.get(subj_title_case, 0) # Use original points dict
                 if base_points > best_sci_base_points:
                      best_sci_base_points = base_points
                      best_sci_name_cuhk = subj_name_lower # Store the lowercase name found

            # Apply FINAL weights
            cuhk_details_list = []
            for subject_name, points in subject_points_dict.items():
                 subj_name_lower = subject_name.lower().strip()
                 final_weight = 1.0
                 if subj_name_lower == CORE_ENG: final_weight = 1.5
                 elif subj_name_lower == CORE_MATH: final_weight = 1.5
                 elif subj_name_lower == best_sci_name_cuhk: final_weight = 2.0
                 current_score = points * final_weight
                 final_weighted_scores_cuhk[subject_name] = current_score # Use original case key
                 # Build details
                 detail_str = f"{subject_name}[{points:.1f}"
                 if final_weight != 1.0: detail_str += f"x{final_weight:.2f}"
                 detail_str += f"={current_score:.2f}]"
                 cuhk_details_list.append(detail_str)

            # Select Best 5 from these final weighted scores (excluding CS)
            sorted_final_scores_cuhk = get_sorted_weighted_scores(final_weighted_scores_cuhk, exclude_cs=True)
            top_5_scores = sorted_final_scores_cuhk[:5]
            final_score = sum(top_5_scores)
            details_msg = f"Method: CUHK Sci Style. Inputs: {', '.join(cuhk_details_list)}."
            details_msg += f" Selected Best {len(top_5_scores)} -> Score: {final_score:.2f}"
            return final_score, details_msg

        # --- Generic Best N Rules ---

        # 6. Best 6 Subjects (Weighted)
        elif "best 6" in scoring_method_clean:
            top_6_scores = sorted_scores_all[:6]
            final_score = sum(top_6_scores)
            details_msg += f" Selected Best {len(top_6_scores)} Weighted -> Score: {final_score:.2f}"
            return final_score, details_msg

        # 7. Best 5 Subjects (Weighted) - Most common fallback
        elif "best 5" in scoring_method_clean:
            top_5_scores = sorted_scores_all[:5]
            final_score = sum(top_5_scores)
            details_msg += f" Selected Best {len(top_5_scores)} Weighted -> Score: {final_score:.2f}"
            return final_score, details_msg

        # 8. Best 4 Subjects (Weighted)
        elif "best 4" in scoring_method_clean:
            top_4_scores = sorted_scores_all[:4]
            final_score = sum(top_4_scores)
            details_msg += f" Selected Best {len(top_4_scores)} Weighted -> Score: {final_score:.2f}"
            return final_score, details_msg

        # 9. 3C+2X (Weighted)
        elif "3c+2x" in scoring_method_clean:
            core_score_sum = 0
            core_details = []
            for core_subj_lower in [CORE_CHI, CORE_ENG, CORE_MATH]:
                score = weighted_scores_dict_lower_keys.get(core_subj_lower, 0)
                core_score_sum += score
                core_details.append(f"{core_subj_lower.title()}={score:.2f}") # Show title case in details

            # Get sorted electives (excl cores, CS)
            other_scores_3c2x = get_sorted_weighted_scores(
                 weighted_scores_dict_lower_keys, exclude_cs=True, exclude_subjects=[CORE_CHI, CORE_ENG, CORE_MATH]
            )
            sum_best_2_electives = sum(other_scores_3c2x[:2])
            final_score = core_score_sum + sum_best_2_electives
            details_msg += f" Selected: 3 Cores ({core_score_sum:.2f}) + Best 2 Electives ({sum_best_2_electives:.2f}) -> Score: {final_score:.2f}"
            return final_score, details_msg

        # --- Fallback: Not Implemented ---
        else:
            details_msg += " Error: Calculation logic not implemented for this method."
            return None, details_msg # Return None score if method not handled

    except Exception as e:
        # Catch potential errors during sorting/summing/etc.
        error_details = initial_details_msg + f" Error during score calculation logic: {e}"
        print(f"--- CALCULATION ERROR ({standardized_id}) ---")
        traceback.print_exc() # Print full traceback for debugging to terminal
        print("-------------------------")
        return None, error_details

# --- Get Student Choices function (should be okay, returns list) ---
def get_student_choices(choices_dataframe, student_regno):
    # ... (Full function code as developed in Part 5B) ...
    if choices_dataframe is None: return None
    global REGNO_COL # Use the potentially updated global REGNO_COL
    regno_column = REGNO_COL
    if regno_column not in choices_dataframe.columns:
        st.error(f"Error: Column '{regno_column}' not found in choices DataFrame.")
        return None
    target_regno_str = str(student_regno).strip()
    try:
        # Ensure comparison is robust, handle potential dtype issues
        matching_rows = choices_dataframe.loc[choices_dataframe[regno_column].astype(str).str.strip() == target_regno_str]
    except Exception as e:
        st.error(f"Error finding student by RegNo: {e}")
        return None

    num_matches = len(matching_rows)
    student_row = None
    if num_matches == 0:
        st.error(f"Error: Student with RegNo '{target_regno_str}' not found.")
        return None
    elif num_matches > 1:
        st.warning(f"Warning: Multiple entries found for RegNo '{target_regno_str}'. Using the first entry.")
        student_row = matching_rows.iloc[0]
    else:
        student_row = matching_rows.iloc[0]
    student_choices_list = []
    for i in range(1, 21):
        choice_col = f"choice{i}" # Use cleaned column name
        if choice_col in student_row.index:
            program_id = student_row[choice_col]
            if pd.notna(program_id) and str(program_id).strip():
                standardized_id = str(program_id).upper().strip()
                student_choices_list.append(standardized_id)
    if not student_choices_list:
        st.warning(f"Student {target_regno_str} found, but no valid choices listed in Choice1-20 columns.")
        # Return empty list instead of None if student found but no choices
        return []
    return student_choices_list


# --- Modify generate_d_day_report to use Streamlit elements ---
def generate_d_day_report(jupas_dataframe, student_choices_dataframe, target_student_regno, actual_student_results):
    """
    Generates a report comparing calculated scores using Streamlit elements.
    Displays header using Class, Class Number, and Name.
    """
    # --- Step 0: Retrieve Student Info using RegNo ---
    # We still need RegNo to reliably find the student row first.
    try:
        regno_col_internal = REGNO_COL # Use global constant
        student_info_row = student_choices_dataframe.loc[student_choices_dataframe[regno_col_internal].astype(str).str.strip() == str(target_student_regno).strip()]
        if student_info_row.empty:
            st.error(f"Internal Error: Could not re-find student info for RegNo '{target_student_regno}'.")
            return
        student_info = student_info_row.iloc[0] # Get the Series for the student

        student_name = student_info.get('name', "N/A")
        student_class = student_info.get('class', "N/A")
        student_class_no = student_info.get('classno', "N/A") # Adjust key if needed

        # --- MODIFIED HEADER ---
        st.header(f"D-Day Report for {student_class} {student_class_no} {student_name}")
        # --- END MODIFICATION ---

    except KeyError:
         st.error(f"Internal Error: Missing required columns (e.g., 'name', 'class', 'classno', '{regno_col_internal}') in student choices data.")
         return
    except Exception as e:
        st.error(f"Error retrieving student details using RegNo: {e}")
        return

    # --- Step 1: Get Student's Choices List (uses RegNo internally, no change needed) ---
    program_choices_list = get_student_choices(student_choices_dataframe, target_student_regno)
    if program_choices_list is None or not program_choices_list:
        return # Error/warning already handled

    # --- Step 1: Get Student's Choices ---
    # Uses st.error/warning internally if needed
    program_choices_list = get_student_choices(student_choices_dataframe, target_student_regno)

    if program_choices_list is None: # Check for None (student not found / error)
        # Error message already shown by get_student_choices
        return
    elif not program_choices_list: # Check for empty list (student found, no choices)
        # Warning message possibly shown by get_student_choices
        return

    st.success(f"Retrieved {len(program_choices_list)} choices for student.")

    # --- Step 2: Process Each Choice ---
    st.subheader("Student's Actual DSE Results:")
    # Use columns for results display
    cols = st.columns(len(actual_student_results))
    i = 0
    for subject, grade in actual_student_results.items():
         with cols[i % len(cols)]: # Cycle through columns
              st.metric(label=subject, value=str(grade))
              i += 1
    # st.json(actual_student_results) # Alternative display


    st.subheader("--- Analysis of Program Choices ---")
    # No need to store in a separate list, process and display directly

    for rank, current_program_id in enumerate(program_choices_list, 1):
        st.divider()
        st.markdown(f"**Choice #{rank}: {current_program_id}**")

        # 2a: Calculate score
        calculated_score, calc_details = calculate_admission_score(jupas_dataframe, current_program_id, actual_student_results)

        # 2b: Retrieve historical scores and name
        lq_score, median_score, uq_score = pd.NA, pd.NA, pd.NA
        prog_name = "N/A"
        if current_program_id in jupas_dataframe.index:
             try:
                 program_row = jupas_dataframe.loc[current_program_id]
                 lq_score = program_row.get('lq_total', pd.NA)
                 median_score = program_row.get('m_total', pd.NA)
                 uq_score = program_row.get('uq_total', pd.NA)
                 prog_name = program_row.get('program_name', "N/A")
             except Exception as e:
                  calc_details += f" (Warning: Error retrieving historical data: {e})"
        else:
             calc_details += f" (Warning: Program ID not found in JUPAS data!)"

        st.markdown(f"*{prog_name[:70]}{'...' if len(prog_name)>70 else ''}*") # Display name italicized
        st.caption(f"Calculation: {calc_details}") # Use caption for details

        # Display scores and suggestion
        score_col, history_col, sugg_col = st.columns([1.5, 2, 2]) # Adjust column widths

        with score_col:
             if calculated_score is not None and pd.notna(calculated_score):
                 st.metric(label="Your Score", value=f"{calculated_score:.2f}")
             else:
                 st.metric(label="Your Score", value="Error")

        with history_col:
             lq_str = f"{lq_score:.2f}" if pd.notna(lq_score) else "N/A"
             med_str = f"{median_score:.2f}" if pd.notna(median_score) else "N/A"
             uq_str = f"{uq_score:.2f}" if pd.notna(uq_score) else "N/A"
             st.markdown(f"**2024 Scores:**\n- LQ: {lq_str}\n- Median: {med_str}\n- UQ: {uq_str}")

        with sugg_col:
             suggestion = ""
             color = "gray" # Default color
             if calculated_score is not None and pd.notna(calculated_score):
                 if pd.notna(lq_score):
                     if calculated_score < lq_score:
                         suggestion = "High Risk (Score < LQ)"
                         color="red"
                     elif pd.notna(median_score) and calculated_score < median_score:
                         suggestion = "Possible/Competitive (LQ ≤ Score < Median)"
                         color="orange"
                     elif pd.notna(uq_score) and calculated_score < uq_score:
                         suggestion = "Good Chance (Median ≤ Score < UQ)"
                         color="blue"
                     elif pd.notna(uq_score) and calculated_score >= uq_score:
                          suggestion = "Likely Safe (Score ≥ UQ)"
                          color="green"
                     elif pd.notna(median_score) and calculated_score >= median_score:
                          suggestion = "Good Chance (Score ≥ Median)"
                          color="blue"
                     else:
                          suggestion = "Possible (Score ≥ LQ)"
                          color="orange"
                 else: # LQ is pd.NA
                     suggestion = "No Comparison Data Available"
                     color="gray"
             else:
                 suggestion = "Calculation Failed"
                 color="gray"

             st.markdown(f"**Suggestion:**")
             st.markdown(f":{color}[{suggestion}]") # Use markdown color


    # Step 4: Print Overall Advice (at the end)
    st.divider()
    st.subheader("--- General Advice ---")
    st.warning("""
    - Review the 'Suggestion' for each program carefully.
    - 'High Risk' programs in Band A/B may warrant reconsideration or reordering.
    - Ensure your top choices are programs you want **and** have a realistic chance for.
    - **Remember:** Admission scores can fluctuate year-to-year. This analysis uses 2024 data as a reference only.
    - Always double-check official university admission requirements and specific program details on the JUPAS website.
    """)


# --- Streamlit App Layout (Main Execution) ---
st.set_page_config(layout="wide")
st.title("JUPAS Calculator & Analysis Tool BETA1 by TTSoultion")

# --- Initialize Session State ---
if 'jupas_data_loaded' not in st.session_state:
    st.session_state.jupas_data_loaded = False
    st.session_state.jupas_df = None
if 'student_choices_loaded' not in st.session_state:
    st.session_state.student_choices_loaded = False
    st.session_state.student_choices_df = None
if 'current_dse_results' not in st.session_state:
    st.session_state.current_dse_results = {} # Store entered DSE results here

# --- Load Main JUPAS Data ---
if not st.session_state.jupas_data_loaded:
    jupas_data = load_and_prepare_jupas_data() # Call the cached function
    if jupas_data is not None:
        st.session_state.jupas_df = jupas_data
        st.session_state.jupas_data_loaded = True
    else:
        st.stop() # Stop the script if main data fails to load

# Get unique institutions list for dropdown (run once)
if 'jupas_df' in st.session_state and st.session_state.jupas_df is not None:
    if 'unique_institutions' not in st.session_state:
         try:
            unique_inst = sorted(st.session_state.jupas_df['institution'].dropna().unique().tolist())
            st.session_state.unique_institutions = unique_inst
            print("Generated unique institutions list for dropdown.") # Log in terminal
         except Exception as e:
             print(f"Could not generate unique institutions list: {e}")
             st.session_state.unique_institutions = [] # Set empty list on error
else:
    # Ensure it's defined even if jupas_df failed loading, though app should stop
     if 'unique_institutions' not in st.session_state:
          st.session_state.unique_institutions = []

# --- Sidebar ---
with st.sidebar:
    st.header("Data Status & Loading")
    # JUPAS status
    if st.session_state.jupas_data_loaded:
        st.success(f"JUPAS Data Loaded ({len(st.session_state.jupas_df)} programs)")
    # Student Choices Loader
    st.divider()
    st.subheader("Load Student Choices")
    uploaded_file = st.file_uploader("Upload CSV File", type=['csv'], key="choices_uploader") # Add key
    # Process upload if file is present
    if uploaded_file is not None:
         # Check if it's a new upload by comparing names or just reload always? Reload always is simpler.
         print(f"Processing uploaded file: {uploaded_file.name}") # Log processing
         choices_data = load_student_choices_from_upload(uploaded_file)
         if choices_data is not None:
             st.session_state.student_choices_df = choices_data
             st.session_state.student_choices_loaded = True
             # Use success message that stays after rerun
             if 'upload_success_msg' not in st.session_state or not st.session_state.upload_success_msg :
                  st.session_state.upload_success_msg = True
                  st.success(f"Choices Loaded ({len(st.session_state.student_choices_df)} records)")
         else:
             st.session_state.student_choices_df = None
             st.session_state.student_choices_loaded = False
             st.session_state.upload_success_msg = False
    elif st.session_state.student_choices_loaded:
         # Keep showing success if already loaded and no new file uploaded
         st.success(f"Choices Loaded ({len(st.session_state.student_choices_df)} records)")
         st.session_state.upload_success_msg = True # Ensure message stays
    else:
         st.warning("Choices file not loaded.")
         st.session_state.upload_success_msg = False
    if st.session_state.student_choices_loaded: # Only show button if data is loaded
        if st.button("Clear Loaded Student Choices", key="clear_choices_btn"):
            st.session_state.student_choices_df = None
            st.session_state.student_choices_loaded = False
            # Clear any related success messages if needed
            if 'upload_success_msg' in st.session_state:
                st.session_state.upload_success_msg = False
            st.info("Cleared loaded student choices.")
            st.rerun() # Rerun to update the UI status

    # DSE Results Status
    st.divider()
    st.subheader("DSE Results Status")
    if st.session_state.current_dse_results:
        st.info(f"{len(st.session_state.current_dse_results)} subjects entered.")
    else:
        st.warning("DSE Results not yet entered.")

# --- Main Page Content (Tabs) ---
st.header("Select Action")
tab1, tab2, tab3, tab4 = st.tabs(["View Program Info", "Enter DSE Results", "Calculate Single Score", "Generate D-Day Report"])

# --- Tab 1: View Program Info ---
with tab1:
    st.subheader("View / Search Program Information")
    # Sub-tabs or radio buttons for search type
    view_option = st.radio("Select Option:", ("View by ID", "Search Name", "Search Institution"), horizontal=True, label_visibility="collapsed", key="tab1_radio")

    jupas_df_display = st.session_state.get('jupas_df', None) # Get from session state

    if jupas_df_display is not None:
        # --- Option 1: View by ID ---
        if view_option == "View by ID":
            prog_id_input = st.text_input("Enter Exact Program ID (e.g., JS1001)", key="view_prog_id")
            if st.button("Show Details by ID", key="view_details_btn"): # Changed button label slightly
                if prog_id_input:
                    # Display details in a dedicated container below the input
                    with st.container(border=True): # Add a border for clarity
                        display_program_details(jupas_df_display, prog_id_input)
                else:
                    st.warning("Please enter a Program ID.")
            # Clear details container if input changes? (More advanced state)

        # --- Option 2: Search Name ---
        elif view_option == "Search Name":
            keyword_input = st.text_input("Enter Keyword(s) for Program Name", key="search_name_kw")
            # Use st.session_state to store search results temporarily
            if 'name_search_results' not in st.session_state:
                 st.session_state.name_search_results = pd.DataFrame() # Initialize empty

            if st.button("Search by Name", key="search_name_btn"):
                if keyword_input:
                    try:
                        keyword_lower = keyword_input.lower()
                        matches = jupas_df_display[jupas_df_display['program_name'].str.contains(keyword_lower, case=False, na=False)]
                        if not matches.empty:
                            st.success(f"Found {len(matches)} program(s) matching '{keyword_input}'.")
                            # Store matches in session state
                            st.session_state.name_search_results = matches[['program_name', 'institution', 'lq_total']] # Store summary
                        else:
                            st.warning(f"No programs found matching '{keyword_input}'.")
                            st.session_state.name_search_results = pd.DataFrame() # Clear results
                    except Exception as e:
                        st.error(f"Error during search: {e}")
                        st.session_state.name_search_results = pd.DataFrame() # Clear results
                else:
                    st.warning("Please enter a keyword.")
                    st.session_state.name_search_results = pd.DataFrame() # Clear results

            # Display results if they exist in session state
            if not st.session_state.name_search_results.empty:
                st.write("Search Results:")
                st.dataframe(st.session_state.name_search_results) # Display summary

                # --- NEW: Input to view details from search results ---
                st.divider()
                detail_id_input = st.text_input("Enter Program ID from results above to view full details:", key="search_name_detail_id")
                if st.button("View Selected Detail", key="search_name_detail_btn"):
                    if detail_id_input:
                         with st.container(border=True):
                              display_program_details(jupas_df_display, detail_id_input)
                    else:
                         st.warning("Please enter a Program ID from the search results.")


        #  --- Option 3: Search Institution ---
        elif view_option == "Search Institution":
            # Get unique list from session state
            institution_options = st.session_state.get('unique_institutions', [])

            if not institution_options:
                 st.warning("Could not retrieve list of institutions.")
                 # Fallback to text input if list generation failed
                 inst_input = st.text_input("Enter Institution Name:", key="search_inst_name_fallback")
                 selected_inst = inst_input if inst_input else None
            else:
                 # Use a dropdown (selectbox)
                 selected_inst = st.selectbox(
                     "Select Institution:",
                     options=[""] + institution_options, # Add blank option at start
                     key="search_inst_select",
                     index=0 # Default to blank
                 )

            # Search button remains the same
            if st.button("Search by Institution", key="search_inst_btn"):
                if selected_inst: # Check if something is selected/entered
                     # Use session state to store results (same key as before)
                     if 'inst_search_results' not in st.session_state:
                          st.session_state.inst_search_results = pd.DataFrame()

                     try:
                        # Use exact match (case-insensitive still good) from dropdown selection
                        # Or contains match if using fallback text input
                        inst_lower = selected_inst.lower()
                        if not institution_options: # If using fallback text input
                            matches = jupas_df_display[jupas_df_display['institution'].str.contains(inst_lower, case=False, na=False)]
                        else: # If using dropdown (exact match preferred, though contains is safer if names have variations)
                            matches = jupas_df_display[jupas_df_display['institution'].str.lower() == inst_lower]
                            # Optional: If exact match fails, try contains?
                            if matches.empty:
                                 matches = jupas_df_display[jupas_df_display['institution'].str.contains(inst_lower, case=False, na=False)]


                        if not matches.empty:
                            st.success(f"Found {len(matches)} program(s) for '{selected_inst}'.")
                            st.session_state.inst_search_results = matches[['program_name', 'scoring_method', 'lq_total']]
                        else:
                            st.warning(f"No programs found for institution '{selected_inst}'.")
                            st.session_state.inst_search_results = pd.DataFrame()
                     except Exception as e:
                         st.error(f"Error during search: {e}")
                         st.session_state.inst_search_results = pd.DataFrame()
                else:
                    st.warning("Please select an institution.")
                    st.session_state.inst_search_results = pd.DataFrame() # Clear results

            # Display results and add detail view input (same logic as before)
            if 'inst_search_results' in st.session_state and not st.session_state.inst_search_results.empty:
                st.write("Search Results:")
                st.dataframe(st.session_state.inst_search_results)
                st.divider()
                detail_id_input_inst = st.text_input("Enter Program ID from results above to view full details:", key="search_inst_detail_id")
                if st.button("View Selected Detail", key="search_inst_detail_btn"):
                    if detail_id_input_inst:
                         with st.container(border=True):
                              display_program_details(jupas_df_display, detail_id_input_inst)
                    else:
                         st.warning("Please enter a Program ID.")


    else:
         st.error("JUPAS data not loaded. Cannot display program info.")


# --- Tab 2: Enter DSE Results ---
# --- Define Grade Options --- (Keep as before)
GRADE_OPTIONS = ["", "5**", "5*", "5", "4", "3", "2", "1", "U"]
CS_GRADE_OPTIONS = ["", "A", "U"] # Specific options for CS

# --- Define Fixed Elective Subject List for User Selection ---
# Using Title Case for user display
FIXED_ELECTIVE_OPTIONS = [
    "", # Blank option first
    "BAFS", # Business, Accounting and Financial Studies
    "Biology",
    "Chemistry",
    "Chinese History",
    "Chinese Literature",
    "Economics",
    "Geography",
    "History", # Added History
    "ICT", # Information and Communication Technology
    "Literature in English", # Added Eng Lit
    "Mathematics M1",
    "Mathematics M2",
    "Music", # Added Music
    "Physical Education", # Added PE
    "Physics",
    "Tourism and Hospitality Studies",
    "Visual Arts"
    # Add any other common electives you want to explicitly list
]
with tab2:
    st.subheader("Enter/Update DSE Results")
    st.caption("Use the dropdown menus to select grades (e.g., 5**, A). Saved results are used for calculations.")

    if 'current_dse_results' not in st.session_state:
        st.session_state.current_dse_results = {}

    # Display Current Results (keep as before)
    if st.session_state.current_dse_results:
        st.write("**Currently Saved Results:**")
        st.json(st.session_state.current_dse_results)
        # --- Clear DSE Button ---
        if st.button("Clear All Entered DSE Results", key="clear_dse_btn"):
            st.session_state.current_dse_results = {}
            st.success("Cleared DSE results.")
            st.rerun() # Rerun to update display and form
        st.divider()
    else:
        st.info("No DSE results entered yet.")

    with st.form(key="dse_form_fixed_electives"):
        st.markdown("**Enter/Edit Grades:**")
        col1, col2 = st.columns(2)

        # --- Core Subjects (Keep as before) ---
        with col1:
            st.markdown("**Core Subjects**")
            core_subjects_grades = {}
            core_subjects_list = ["Chinese Language", "English Language", "Mathematics Compulsory Part"]
            cs_subject = "Citizenship and Social Development"

            for subj in core_subjects_list:
                 current_grade = st.session_state.current_dse_results.get(subj, "")
                 try: current_index = GRADE_OPTIONS.index(current_grade)
                 except ValueError: current_index = 0
                 core_subjects_grades[subj] = st.selectbox(f"{subj}:", options=GRADE_OPTIONS, key=f"dse_core_{subj}", index=current_index)

            current_cs_grade = st.session_state.current_dse_results.get(cs_subject, "A")
            try: current_cs_index = CS_GRADE_OPTIONS.index(current_cs_grade)
            except ValueError: current_cs_index = CS_GRADE_OPTIONS.index("A")
            core_subjects_grades[cs_subject] = st.selectbox(f"{cs_subject}:", options=CS_GRADE_OPTIONS, key=f"dse_core_{cs_subject}", index=current_cs_index)


        # --- Elective Subjects (Using FIXED_ELECTIVE_OPTIONS) ---
        with col2:
            st.markdown("**Elective Subjects**")
            num_elective_slots = 5 # Still use fixed slots, but dropdown is fixed list
            elective_subjects_grades = {}

            # Try to pre-populate slots based on current results
            core_and_cs = core_subjects_list + [cs_subject]
            current_electives = {k:v for k,v in st.session_state.current_dse_results.items() if k not in core_and_cs}
            current_elective_pairs = list(current_electives.items())

            for i in range(num_elective_slots):
                default_subj = ""
                default_grade = ""
                if i < len(current_elective_pairs):
                    # Match stored subject name (which should be standard) to fixed list
                    stored_subj = current_elective_pairs[i][0]
                    if stored_subj in FIXED_ELECTIVE_OPTIONS: # Check if stored subject is in our fixed list
                         default_subj = stored_subj
                         default_grade = current_elective_pairs[i][1]
                    # Else: stored subject isn't in fixed list, leave default blank

                # Find index for subject and grade selectboxes
                try: subj_index = FIXED_ELECTIVE_OPTIONS.index(default_subj)
                except ValueError: subj_index = 0 # Default to blank if not found
                try: grade_index = GRADE_OPTIONS.index(default_grade)
                except ValueError: grade_index = 0

                st.markdown(f"Elective {i+1}", help="Select subject AND grade")
                # Use the FIXED list for options now
                selected_subj = st.selectbox(
                    f"Subject {i+1}:",
                    options=FIXED_ELECTIVE_OPTIONS, # Use the fixed list
                    key=f"dse_elec_subj_{i}",
                    index=subj_index
                )
                selected_grade = st.selectbox(
                    f"Grade {i+1}:",
                    options=GRADE_OPTIONS,
                    key=f"dse_elec_grade_{i}",
                    index=grade_index
                 )

                if selected_subj: # Only store if a subject name is chosen
                    elective_subjects_grades[selected_subj] = selected_grade


        # --- Submit Button (Keep as before) ---
        submitted = st.form_submit_button("Save / Update All DSE Results")
        if submitted:
            final_results = {}
            # Add core subjects if grade selected
            for subj, grade in core_subjects_grades.items():
                if grade: final_results[subj] = grade

            # Add electives if both subject and grade selected
            processed_electives = set() # Keep track of subjects added
            for subj, grade in elective_subjects_grades.items():
                 if subj and grade: # Check both subject and grade are not blank
                    if subj not in processed_electives:
                         final_results[subj] = grade
                         processed_electives.add(subj)
                    else:
                         # Only warn once per duplicate subject in the form submission
                         if subj not in getattr(st, f"warned_{subj}", []): # Use temporary attr to track warning
                             st.warning(f"Duplicate elective subject '{subj}' selected in form. Using first entry.")
                             setattr(st, f"warned_{subj}", True) # Mark as warned for this run
            # Update session state
            st.session_state.current_dse_results = final_results
            st.success("DSE Results Updated!")
            st.rerun() # Rerun to show updated results


    # Display current results outside the form too (Keep as before)
    st.divider()
    if st.session_state.current_dse_results:
        st.write("**Current Stored DSE Results (Live):**")
        st.json(st.session_state.current_dse_results)

# --- Tab 3: Calculate Single Score ---
with tab3:
    st.subheader("Calculate Estimated Score for a Single Program")
    st.caption("Uses the DSE results entered in the 'Enter DSE Results' tab.")

    # Initialize session state for storing calculation results for this tab
    if 'calc_score_result' not in st.session_state:
        st.session_state.calc_score_result = None # Will store tuple (score, details) or None
    if 'calc_prog_id_processed' not in st.session_state:
        st.session_state.calc_prog_id_processed = None # Store the ID that was processed


    # Check if DSE results have been entered
    if not st.session_state.get('current_dse_results', {}):
        st.warning("⚠️ Please enter DSE results in the 'Enter DSE Results' tab first.")
        # Clear previous results if DSE results are cleared
        st.session_state.calc_score_result = None
        st.session_state.calc_prog_id_processed = None
    else:
        jupas_df_calc = st.session_state.get('jupas_df', None)
        if jupas_df_calc is None:
            st.error("❌ JUPAS program data is not loaded. Cannot calculate score.")
            st.session_state.calc_score_result = None
            st.session_state.calc_prog_id_processed = None
        else:
            # Input for Program ID
            prog_id_input_calc = st.text_input("Enter Program ID to Calculate:", key="calc_prog_id").strip().upper()

            # Button to trigger calculation
            if st.button("Calculate Score Now", key="calc_score_btn"):
                if not prog_id_input_calc:
                    st.warning("⚠️ Please enter a Program ID.")
                    st.session_state.calc_score_result = None # Clear previous results
                    st.session_state.calc_prog_id_processed = None
                else:
                    # Perform calculation and store result in session state
                    with st.spinner(f"Calculating score for {prog_id_input_calc}..."):
                        score, details = calculate_admission_score(
                            jupas_df_calc,
                            prog_id_input_calc,
                            st.session_state.current_dse_results
                        )
                    # Store the outcome (score might be None)
                    st.session_state.calc_score_result = (score, details)
                    st.session_state.calc_prog_id_processed = prog_id_input_calc
                    # No st.rerun() needed here, display happens below based on state

    # --- Display Area: Only shows results based on session state ---
    st.markdown("---") # Divider
    # Check if a calculation was performed for a specific ID
    if st.session_state.calc_prog_id_processed:
        calculated_id = st.session_state.calc_prog_id_processed
        result_tuple = st.session_state.calc_score_result

        if result_tuple: # Check if result tuple exists
            score, details = result_tuple

            # --- Case 1: Calculation Succeeded ---
            if score is not None and pd.notna(score):
                st.metric(label=f"Estimated Score for {calculated_id}", value=f"{score:.2f}")
                st.caption(f"Calculation Details: {details}")

                # Add Comparison to Historical Data
                st.markdown("**Comparison with 2024 Data:**")
                try:
                    jupas_df_display = st.session_state.jupas_df # Use loaded df
                    program_row = jupas_df_display.loc[calculated_id]
                    lq = program_row.get('lq_total', pd.NA)
                    median = program_row.get('m_total', pd.NA)
                    uq = program_row.get('uq_total', pd.NA)

                    col1, col2, col3 = st.columns(3)
                    lq_str = f"{lq:.2f}" if pd.notna(lq) else "N/A"
                    med_str = f"{median:.2f}" if pd.notna(median) else "N/A"
                    uq_str = f"{uq:.2f}" if pd.notna(uq) else "N/A"
                    col1.metric("Lower Quartile", lq_str)
                    col2.metric("Median", med_str)
                    col3.metric("Upper Quartile", uq_str)

                    # Generate suggestion text, emoji, color
                    suggestion_text, suggestion_emoji, bg_color = "", "", "#f0f2f6" # Defaults
                    color_map = {"red": "#FFCCCC", "orange": "#FFE5CC", "blue": "#CCE5FF", "green": "#CCFFCC", "gray": "#f0f2f6"}

                    if pd.notna(lq):
                        if score < lq:
                            suggestion_text, suggestion_emoji, bg_color = "High Risk (Score < LQ)", "🔴", color_map["red"]
                        elif pd.notna(median) and score < median:
                            suggestion_text, suggestion_emoji, bg_color = "Possible/Competitive (LQ ≤ Score < Median)", "🟠", color_map["orange"]
                        elif pd.notna(uq) and score < uq:
                            suggestion_text, suggestion_emoji, bg_color = "Good Chance (Median ≤ Score < UQ)", "🔵", color_map["blue"]
                        elif pd.notna(uq) and score >= uq:
                             suggestion_text, suggestion_emoji, bg_color = "Likely Safe (Score ≥ UQ)", "🟢", color_map["green"]
                        elif pd.notna(median) and score >= median:
                             suggestion_text, suggestion_emoji, bg_color = "Good Chance (Score ≥ Median)", "🔵", color_map["blue"]
                        else:
                             suggestion_text, suggestion_emoji, bg_color = "Possible (Score ≥ LQ)", "🟠", color_map["orange"]
                    else:
                        suggestion_text, suggestion_emoji, bg_color = "No historical data for comparison.", "⚪", color_map["gray"]

                    # Display Suggestion Box
                    st.markdown(f"**Suggestion:**")
                    styled_suggestion_html = f"""
                    <div style="background-color:{bg_color}; border:1px solid #cccccc; padding:10px; border-radius:5px; margin-top:10px;">
                        <span style="font-size: 1.1em; font-weight: bold;">{suggestion_emoji} {suggestion_text}</span>
                    </div>"""
                    st.markdown(styled_suggestion_html, unsafe_allow_html=True)

                except KeyError:
                    st.warning(f"Historical score data not found for {calculated_id}.")
                except Exception as e:
                    st.error(f"Error retrieving/comparing historical scores: {e}")

            # --- Case 2: Calculation Failed ---
            else:
                st.error(f"Could not calculate score for {calculated_id}.")
                st.caption(f"Reason: {details}") # Show the error/details message
        else:
            # Should not happen if calc_prog_id_processed is set, but good fallback
             st.info("Enter a Program ID and click 'Calculate Score Now'.")


# --- Tab 4: Generate D-Day Report ---
with tab4:
    st.subheader("Generate D-Day Report for Student")

    # --- Prerequisite Checks (Keep as before) ---
    prereqs_met = True
    # ... (checks for student_choices_df, current_dse_results, jupas_df) ...

    # --- Inputs for Class and Class Number ---
    if prereqs_met:
        st.markdown("**Select Student:**")
        # Define fixed class options directly here or reference global constant
        class_options_fixed = ["", "6A", "6B", "6C", "6D"] # Or use FIXED_CLASS_OPTIONS

        col1, col2 = st.columns(2)
        with col1:
            # Use the fixed list directly in the selectbox
            selected_class = st.selectbox(
                "Select Class:",
                options=class_options_fixed, # Use the fixed list
                key="report_class_select_fixed",
                index=0 # Default blank
            )
        with col2:
            entered_class_no = st.text_input("Enter Class Number:", key="report_classno").strip()

        # --- Generate Button (Logic remains the same) ---
        if st.button("Generate D-Day Report", key="gen_report_btn_class"):
            if not selected_class or not entered_class_no:
                st.warning("⚠️ Please select a Class and enter a Class Number.")
            else:
                # --- Find the student (Logic remains the same) ---
                student_choices_df = st.session_state.student_choices_df
                jupas_df_report = st.session_state.jupas_df
                dse_results_report = st.session_state.current_dse_results
                class_col = 'class' # Adjust if needed
                classno_col = 'classno' # Adjust if needed

                if class_col not in student_choices_df.columns or classno_col not in student_choices_df.columns:
                     st.error(f"Error: Required columns ('{class_col}', '{classno_col}') not found in choices data.")
                else:
                    try:
                        # Filter based on BOTH class and class number
                        # Ensure comparison handles potential case differences in data if needed
                        matches = student_choices_df[
                            (student_choices_df[class_col].astype(str).str.strip().str.fullmatch(selected_class, case=False)) & # Use fullmatch case-insensitive
                            (student_choices_df[classno_col].astype(str).str.strip() == entered_class_no)
                        ]

                        # ...(Rest of the logic to handle matches=0, >1, or =1 and call generate_d_day_report remains the same)...
                        if len(matches) == 0:
                            st.error(f"❌ Student not found for Class '{selected_class}' and Class Number '{entered_class_no}'.")
                        # ... etc ...
                        else: # Found one or more, use first
                            regno_col_report = REGNO_COL
                            if regno_col_report not in matches.columns:
                                st.error(f"Cannot proceed: RegNo column '{regno_col_report}' missing.")
                            else:
                                if len(matches) > 1:
                                     st.warning(f"⚠️ Multiple students found for {selected_class} {entered_class_no}. Using first match.")
                                target_regno = matches.iloc[0][regno_col_report]
                                with st.spinner(f"Generating report for {selected_class} {entered_class_no}..."):
                                    generate_d_day_report(jupas_df_report, student_choices_df, target_regno, dse_results_report)
                    except Exception as e:
                         st.error(f"An error occurred while finding the student: {e}")