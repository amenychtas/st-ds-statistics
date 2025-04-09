# Required libraries
import streamlit as st
import pandas as pd
import io # Re-added for BytesIO
import math # To help calculate number of columns for checkboxes

# --- Streamlit App Configuration ---
# Set the page title and layout
st.set_page_config(page_title="Excel File Merger & Analyzer", layout="wide")

# --- Helper Function to Convert DataFrame to Excel Bytes ---
# Re-added function to handle Excel conversion for downloads
@st.cache_data
def to_excel(df):
    """
    Converts a Pandas DataFrame to an Excel file format in memory (bytes).

    Args:
        df (pd.DataFrame): The DataFrame to convert.

    Returns:
        bytes: The Excel file content as bytes.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# --- Main Application Logic ---

# --- Title and Description ---
st.title("📊 Excel File Merger & Analyzer")
st.write("""
Upload multiple Excel files (.xlsx). The first row of each file will be skipped.
The files will then be merged. Columns will be renamed ('Περίοδος δήλωσης'->'Περίοδος', 'Τμήμα Τάξης'->'Μάθημα', 'Αριθμός Μητρώου'->'Έτος Εγγραφής').
Only the columns 'Περίοδος', 'Μάθημα', 'Έτος Εγγραφής', 'Βαθμολογία' will be kept (in that order),
and the 'Έτος Εγγραφής' will be truncated to the first 3 characters.
Select 'Περίοδος' to view overall summary by 'Μάθημα'.
Then, select 'Μάθημα' to view summary by 'Έτος Εγγραφής' for the selected periods and courses.
You can download each summary table as an XLSX file.
""") # Updated description

# --- Column Renaming and Definitions ---
# Define original column names expected from Excel
ORIGINAL_PERIODOS = "Περίοδος δήλωσης"
ORIGINAL_TMIMA = "Τμήμα Τάξης"
ORIGINAL_MITROO = "Αριθμός Μητρώου"
ORIGINAL_GRADE = "Βαθμολογία"

# Define new column names
PERIODOS_COLUMN = "Περίοδος"
TMIMA_COLUMN = "Μάθημα"
MITROO_COLUMN = "Έτος Εγγραφής"
GRADE_COLUMN = "Βαθμολογία" # Remains the same

# Define the columns to keep after renaming and their order
FIXED_COLUMNS = [PERIODOS_COLUMN, TMIMA_COLUMN, MITROO_COLUMN, GRADE_COLUMN]

# Renaming dictionary
RENAME_DICT = {
    ORIGINAL_PERIODOS: PERIODOS_COLUMN,
    ORIGINAL_TMIMA: TMIMA_COLUMN,
    ORIGINAL_MITROO: MITROO_COLUMN
    # GRADE_COLUMN remains the same, no need to include
}


# --- State Management Initialization ---
# Use Streamlit's session state to store data across reruns.
# Initialize state variables if they don't exist.
if 'combined_df' not in st.session_state:
    st.session_state.combined_df = None # To store the raw merged DataFrame (before renaming)
if 'processed_df' not in st.session_state:
    st.session_state.processed_df = None # To store the final DataFrame after renaming, selection and modification
# Aggregations are now calculated dynamically based on selections

# --- File Uploader ---
# Create a file uploader widget that accepts multiple .xlsx files.
uploaded_files = st.file_uploader(
    "Choose your Excel files",
    type="xlsx",  # Restrict file type to .xlsx
    accept_multiple_files=True, # Allow uploading more than one file
    help="Upload one or more .xlsx files. The first row will be skipped."
)

# --- Processing Uploaded Files ---
# This block executes only if files have been uploaded *in the current run*
# AND if the processed_df hasn't been generated yet in this session.
if uploaded_files and st.session_state.processed_df is None:
    # Clear previous state if new files are uploaded implicitly resetting
    st.session_state.combined_df = None
    st.session_state.processed_df = None
    to_excel.clear() # Clear cache if new files are uploaded

    dataframes = [] # List to hold DataFrames from each file
    file_names = [file.name for file in uploaded_files] # Get names for display
    st.write(f"Processing files: {', '.join(file_names)}")

    # Show a progress bar while processing files
    progress_bar = st.progress(0)
    status_text = st.empty() # Placeholder for status messages
    error_occurred = False # Flag to track errors

    try:
        # Loop through each uploaded file
        for i, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            status_text.text(f"Reading {file_name} (skipping first row)...")
            try:
                # Read the current Excel file into a pandas DataFrame
                uploaded_file.seek(0) # Reset file pointer
                df = pd.read_excel(uploaded_file, header=0, skiprows=1)

                # Basic check if expected original columns exist before proceeding
                # Note: This assumes all files should have these columns. Adjust if needed.
                required_original_cols = [ORIGINAL_PERIODOS, ORIGINAL_TMIMA, ORIGINAL_MITROO, ORIGINAL_GRADE]
                if not all(col in df.columns for col in required_original_cols):
                     st.warning(f"⚠️ File '{file_name}' is missing one or more expected original columns ({', '.join(required_original_cols)}). Skipping this file.")
                     continue # Skip to the next file

                if not df.empty:
                    dataframes.append(df)
                else:
                    st.warning(f"⚠️ No data found in '{file_name}' after skipping the first row.")

            except Exception as file_error:
                st.error(f"❌ Error processing file '{file_name}': {file_error}")
                error_occurred = True # Set flag if any file fails

            # Update the progress bar
            progress_bar.progress((i + 1) / len(uploaded_files))

        status_text.text("Combining files...")
        if dataframes and not error_occurred:
            # Concatenate all DataFrames
            combined_df_temp = pd.concat(dataframes, ignore_index=True)
            st.session_state.combined_df = combined_df_temp # Store pre-renamed combined df
            st.success("✅ Files successfully merged (first row skipped)!")

            # --- Apply Renaming ---
            status_text.text("Renaming columns...")
            combined_df_temp.rename(columns=RENAME_DICT, inplace=True)
            st.success("✅ Columns renamed successfully!")

            # --- Apply Fixed Selection and Modification ---
            status_text.text("Applying column selection and processing...")
            # Check for NEW column names
            missing_cols = [col for col in FIXED_COLUMNS if col not in combined_df_temp.columns]

            if missing_cols:
                # Display error with NEW column names
                st.error(f"❌ Error: The following required columns are missing after renaming: {', '.join(missing_cols)}")
                st.session_state.processed_df = None
                error_occurred = True
            else:
                # Select columns in fixed order using NEW names and create a copy
                processed_df = combined_df_temp[FIXED_COLUMNS].copy()

                # Modify the NEW "Έτος Εγγραφής" column
                processed_df[MITROO_COLUMN] = processed_df[MITROO_COLUMN].astype(str).str[:3]

                # Ensure GRADE_COLUMN is numeric for later aggregation
                processed_df[GRADE_COLUMN] = pd.to_numeric(processed_df[GRADE_COLUMN], errors='coerce')

                # Store the final processed DataFrame in session state
                st.session_state.processed_df = processed_df
                st.success("✅ Column selection and processing applied successfully!")
                # Aggregations are now performed later based on selections


        elif not dataframes and not error_occurred:
             st.warning("⚠️ No data was extracted from the uploaded files after skipping the first row.")
             st.session_state.combined_df = None
             st.session_state.processed_df = None
        elif error_occurred:
             st.error("Processing stopped due to errors. Please check the messages above.")
             st.session_state.processed_df = None

        # Clear progress bar and status text after completion or error
        progress_bar.empty()
        status_text.empty()

    except Exception as e:
        # Catch unexpected errors during the overall process
        st.error(f"❌ An unexpected error occurred during file processing: {e}")
        # Clear all state on major error
        st.session_state.combined_df = None
        st.session_state.processed_df = None
        progress_bar.empty() # Ensure progress bar is cleared
        status_text.empty() # Ensure status text is cleared
        error_occurred = True


# --- Display Processed Data ---
# Display preview of the final processed data
if st.session_state.processed_df is not None:
    st.subheader("Processed Data Preview (Renamed Columns)") # Updated subheader
    st.dataframe(st.session_state.processed_df)

    # --- Step 1: Select Periodos ---
    st.subheader(f"Βήμα 1: Επιλογή {PERIODOS_COLUMN} για Ανάλυση") # Use new name
    processed_df = st.session_state.processed_df
    # Get options from the NEW column name
    periodos_options = sorted(processed_df[PERIODOS_COLUMN].unique().tolist())
    selected_periodoi = []

    st.write(f"Επιλέξτε μία ή περισσότερες **{PERIODOS_COLUMN}**:") # Use new name
    # Use columns for better layout
    num_periodos_options = len(periodos_options)
    cols_per_row_p = 4 # Adjust number of columns as needed
    num_rows_p = math.ceil(num_periodos_options / cols_per_row_p)

    option_index_p = 0
    for _ in range(num_rows_p):
        cols_p = st.columns(cols_per_row_p)
        for j in range(cols_per_row_p):
            if option_index_p < num_periodos_options:
                periodos = periodos_options[option_index_p]
                checkbox_key_p = f"periodos_cb_{periodos}" # Key uses value, not name
                is_checked_p = cols_p[j].checkbox(periodos, key=checkbox_key_p)
                if is_checked_p:
                    selected_periodoi.append(periodos)
                option_index_p += 1

    # Proceed only if at least one Periodos is selected
    if selected_periodoi:
        # Filter data based on selected Periodos (using NEW column name)
        df_filtered_by_periodos = processed_df[processed_df[PERIODOS_COLUMN].isin(selected_periodoi)]

        if not df_filtered_by_periodos.empty:
            # --- Display Overall Aggregation by Tmima Taksis (now "Μάθημα") ---
            st.subheader(f"Συνολική Σύνοψη ανά {TMIMA_COLUMN} (για Επιλεγμένες Περιόδους)") # Use new name
            # Define aggregation logic using NEW column names and NEW aggregation names
            aggregation_logic = {
                'Εγγεγραμμένοι': (TMIMA_COLUMN, 'size'), # Updated Name
                'Συμμετείχαν': (GRADE_COLUMN, 'count'),   # Updated Name
                'Επιτυχόντες': (GRADE_COLUMN, lambda x: (x >= 5).sum()) # Updated Name
            }
            # Group by NEW column name
            tmima_agg = df_filtered_by_periodos.groupby(TMIMA_COLUMN).agg(**aggregation_logic).reset_index()
            st.dataframe(tmima_agg)

            # --- Add Download Button for Tmima Aggregation ---
            tmima_excel_data = to_excel(tmima_agg)
            st.download_button(
                label=f"📥 Λήψη Σύνοψης ανά {TMIMA_COLUMN} (.xlsx)", # Use new name
                data=tmima_excel_data,
                file_name=f"summary_by_{TMIMA_COLUMN}.xlsx", # Use new name in filename
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help=f"Λήψη της σύνοψης ανά {TMIMA_COLUMN} για τις επιλεγμένες περιόδους." # Use new name
            )
            # --- End Download Button ---

            # --- Step 2: Select Tmima Taksis (now "Μάθημα") ---
            st.subheader(f"Βήμα 2: Επιλογή {TMIMA_COLUMN} για Ανάλυση ανά {MITROO_COLUMN}") # Use new names
            # Get options from the NEW column name in the filtered data
            tmima_options = sorted(df_filtered_by_periodos[TMIMA_COLUMN].unique().tolist())
            selected_tmimata = []

            st.write(f"Επιλέξτε ένα ή περισσότερα **{TMIMA_COLUMN}** από τα παραπάνω:") # Use new name
            # Use columns for better layout
            num_tmima_options = len(tmima_options)
            cols_per_row_t = 4 # Adjust number of columns as needed
            num_rows_t = math.ceil(num_tmima_options / cols_per_row_t)

            option_index_t = 0
            for _ in range(num_rows_t):
                cols_t = st.columns(cols_per_row_t)
                for j in range(cols_per_row_t):
                    if option_index_t < num_tmima_options:
                        tmima = tmima_options[option_index_t]
                        checkbox_key_t = f"tmima_cb_{tmima}" # Key uses value
                        is_checked_t = cols_t[j].checkbox(tmima, key=checkbox_key_t)
                        if is_checked_t:
                            selected_tmimata.append(tmima)
                        option_index_t += 1

            # --- Perform and display aggregation by Mitroo (now "Έτος Εγγραφής") ---
            if selected_tmimata:
                selected_periodoi_str = ", ".join(selected_periodoi)
                selected_tmimata_str = ", ".join(selected_tmimata)
                # Use new names in the message
                st.write(f"Εμφάνιση σύνοψης ανά {MITROO_COLUMN} για Περιόδους: **{selected_periodoi_str}** και Μαθήματα: **{selected_tmimata_str}**")

                # Filter data further based on selected Tmimata (using NEW column name)
                df_filtered_by_both = df_filtered_by_periodos[df_filtered_by_periodos[TMIMA_COLUMN].isin(selected_tmimata)]

                if not df_filtered_by_both.empty:
                    # Group by the NEW "Έτος Εγγραφής" column using updated aggregation logic
                    mitroo_agg_filtered = df_filtered_by_both.groupby(MITROO_COLUMN).agg(**aggregation_logic).reset_index()
                    st.dataframe(mitroo_agg_filtered)

                    # --- Add Download Button for Mitroo Aggregation ---
                    mitroo_excel_data = to_excel(mitroo_agg_filtered)
                    st.download_button(
                        label=f"📥 Λήψη Σύνοψης ανά {MITROO_COLUMN} (.xlsx)", # Use new name
                        data=mitroo_excel_data,
                        file_name=f"summary_by_{MITROO_COLUMN}_filtered.xlsx", # Use new name in filename
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help=f"Λήψη της σύνοψης ανά {MITROO_COLUMN} για τις επιλεγμένες περιόδους και μαθήματα." # Use new name
                    )
                    # --- End Download Button ---

                else:
                    st.info(f"Δεν βρέθηκαν δεδομένα για τον συνδυασμό των επιλεγμένων περιόδων και μαθημάτων.") # Updated text
            else:
                 st.info(f"Παρακαλώ επιλέξτε (τικάρετε) ένα ή περισσότερα {TMIMA_COLUMN} παραπάνω για να δείτε την ανάλυση ανά {MITROO_COLUMN}.") # Use new name

        else:
            st.info(f"Δεν βρέθηκαν δεδομένα για τις επιλεγμένες {PERIODOS_COLUMN}.") # Use new name
    else:
        st.info(f"Παρακαλώ επιλέξτε (τικάρετε) μία ή περισσότερες {PERIODOS_COLUMN} παραπάνω για να ξεκινήσετε την ανάλυση.") # Use new name


elif uploaded_files:
    # This case handles if processing failed after upload
     st.warning("⚠️ Δεν είναι δυνατή η εμφάνιση της ανάλυσης. Ελέγξτε για σφάλματα επεξεργασίας παραπάνω.")

# --- Footer/Instructions ---
st.markdown("---")
st.caption("App created using Streamlit | Skips first row, merges files, renames & keeps/processes specific columns, and displays summaries based on selections.") # Updated caption
