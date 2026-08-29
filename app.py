import streamlit as st
import pandas as pd
import io

import phonenumbers

def parse_phone(phone_str):
    try:
        phone_str = str(phone_str).strip()
        # If it's a long string of digits without a plus, assume the first digits are country code
        if phone_str.isdigit() and len(phone_str) > 10 and not phone_str.startswith('+'):
            phone_str = '+' + phone_str
            
        # Parse the number, fallback to India ('IN') if no country code provided and 10 digits
        parsed = phonenumbers.parse(phone_str, "IN")
        
        if phonenumbers.is_valid_number(parsed):
            country_code = f"+{parsed.country_code}"
            national_number = str(parsed.national_number)
            return country_code, national_number
        else:
            return None, phone_str
    except Exception:
        return None, phone_str

st.set_page_config(page_title="Data Processing App", layout="wide")

st.title("Bulk Data Separation and Matching App")

tab1, tab2 = st.tabs(["Step 1: 10x Stats DB Processing", "Step 2: Multi-File Matching"])

with tab1:
    st.header("Step 1: Process 10x Stats DB")
    st.markdown("Upload the 10x stats DB file to extract customer details and split phone numbers.")
    
    file_10x = st.file_uploader("Upload 10x Stats DB (CSV or Excel)", type=["csv", "xlsx"], key="10x")
    
    if file_10x:
        try:
            if file_10x.name.endswith(".csv"):
                df_10x = pd.read_csv(file_10x)
            else:
                df_10x = pd.read_excel(file_10x)
                
            st.success(f"Successfully loaded {file_10x.name}")
            st.dataframe(df_10x.head())
            
            st.subheader("Process & Extract")
            st.write("This will extract: `email`, `phoneNumber`, `paymentAmount`, `lms_access`, `half_access`, `intro_call_date` and separate the country code.")
            
            if st.button("Process & Extract Data"):
                required_cols_from_file = [
                    "email", "phoneNumber", "original_phone_number", "customerName", 
                    "paymentAmount", "remainingAmount", "lms_access", 
                    "intro_call_date", "courseName"
                ]
                
                # Check for missing columns and warn
                missing = [col for col in required_cols_from_file if col not in df_10x.columns]
                if missing:
                    st.warning(f"Note: These columns were not found in your file and will be added as empty: {', '.join(missing)}")
                
                # Create DataFrame with ONLY the required columns
                processed_df = pd.DataFrame()
                for col in required_cols_from_file:
                    if col in df_10x.columns:
                        processed_df[col] = df_10x[col]
                    else:
                        processed_df[col] = None
                        
                # Split phone numbers if phoneNumber exists
                if "phoneNumber" in processed_df.columns:
                    country_codes = []
                    local_numbers = []
                    for phone in processed_df["phoneNumber"]:
                        cc, loc = parse_phone(phone)
                        country_codes.append(cc)
                        local_numbers.append(loc)
                        
                    # Insert the country code right before phoneNumber
                    loc_idx = processed_df.columns.get_loc("phoneNumber")
                    processed_df.insert(loc_idx, "Country Code", country_codes)
                    processed_df["phoneNumber"] = local_numbers
                    
                    # Convert to string to prevent Excel from converting to scientific notation
                    processed_df["Country Code"] = processed_df["Country Code"].astype(str)
                    processed_df["phoneNumber"] = processed_df["phoneNumber"].astype(str)
                
                # Reorder columns to exactly match the requested output
                final_order = [
                    "email", "Country Code", "phoneNumber", "original_phone_number", 
                    "customerName", "paymentAmount", "remainingAmount", 
                    "lms_access", "intro_call_date", "courseName"
                ]
                # Only include columns that actually exist (in case phoneNumber was missing and Country Code wasn't created)
                final_order = [c for c in final_order if c in processed_df.columns]
                processed_df = processed_df[final_order]
                    
                # Sometimes Excel still forces numbers to scientific notation if they look like numbers.
                # Adding a zero-width space or a normal space at the front can prevent this if needed, 
                # but usually astype(str) is enough for .xlsx format.
                
                st.success("Data processed successfully!")
                st.dataframe(processed_df.head())
                
                # Download button
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    processed_df.to_excel(writer, index=False, sheet_name='Filtered_Data')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="Download Filtered Data as Excel",
                    data=processed_data,
                    file_name="Step1_Filtered_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Error reading file: {e}")

with tab2:
    st.header("Step 2: Match Data Across Multiple Files")
    st.markdown("Upload multiple files to merge them based on Email ID and Phone Number.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        file_filtered = st.file_uploader("Upload Filtered 10x Data (from Step 1)", type=["csv", "xlsx"], key="filtered")
        file_main = st.file_uploader("Upload Main Data", type=["csv", "xlsx"], key="main")
        file_lms = st.file_uploader("Upload LMS Export Batch Data", type=["csv", "xlsx"], key="lms")
        
    with col2:
        file_pref = st.file_uploader("Upload Batch Preference File", type=["csv", "xlsx"], key="pref")
        file_agent = st.file_uploader("Upload Agent Batch Change Data", type=["csv", "xlsx"], key="agent")
        file_process = st.file_uploader("Upload Data Process File", type=["csv", "xlsx"], key="process")
        
    all_files = [file_filtered, file_main, file_lms, file_pref, file_agent, file_process]
    
    if any(all_files):
        st.subheader("Uploaded Files Summary")
        for f in all_files:
            if f:
                st.write(f"- {f.name}")
                
        if st.button("Merge & Match Data"):
            if not file_filtered:
                st.error("Please upload the Filtered 10x Data (from Step 1) as it is required as the base file.")
            else:
                try:
                    # Helper function for FULL OUTER OR matching (Email first, then Phone)
                    def merge_datasets(base_df, right_df, email_col_right, phone_col_right):
                        if email_col_right and email_col_right in right_df.columns:
                            right_df["_match_email"] = right_df[email_col_right].astype(str).str.lower().str.strip()
                        else:
                            right_df["_match_email"] = None
                            
                        if phone_col_right and phone_col_right in right_df.columns:
                            right_df["_match_phone"] = right_df[phone_col_right].astype(str).str.replace(r'\D', '', regex=True)
                        else:
                            right_df["_match_phone"] = None
                            
                        cols_to_add = [c for c in right_df.columns if c not in ["_match_email", "_match_phone"]]
                        
                        # Ensure base_df has keys
                        if "email_lower" not in base_df.columns: base_df["email_lower"] = ""
                        if "phone_str" not in base_df.columns: base_df["phone_str"] = ""
                        
                        # Create new columns in base_df
                        col_mapping = {}
                        for col in cols_to_add:
                            final_col = col
                            while final_col in base_df.columns:
                                final_col = final_col + "_dup"
                            col_mapping[col] = final_col
                            base_df[final_col] = None
                            
                        # Build index of base_df for fast lookups
                        base_email_to_idx = {}
                        base_phone_to_idx = {}
                        for idx, row in base_df.iterrows():
                            em = str(row.get("email_lower", "")).lower().strip()
                            ph = str(row.get("phone_str", "")).strip()
                            if em and em not in ['nan', 'none', '']: base_email_to_idx[em] = idx
                            if ph and ph not in ['nan', 'none', '']: base_phone_to_idx[ph] = idx
                            
                        unmatched_rows = []
                        
                        for _, right_row in right_df.iterrows():
                            r_email = str(right_row.get("_match_email", "")).strip()
                            r_phone = str(right_row.get("_match_phone", "")).strip()
                            
                            match_idx = None
                            if r_email and r_email not in ['nan', 'none'] and r_email in base_email_to_idx:
                                match_idx = base_email_to_idx[r_email]
                            elif r_phone and r_phone not in ['nan', 'none'] and r_phone in base_phone_to_idx:
                                match_idx = base_phone_to_idx[r_phone]
                                
                            if match_idx is not None:
                                # Update existing row
                                for col in cols_to_add:
                                    base_df.at[match_idx, col_mapping[col]] = right_row[col]
                            else:
                                # Append as new row
                                new_row = {c: None for c in base_df.columns}
                                new_row["email_lower"] = r_email
                                new_row["phone_str"] = r_phone
                                
                                # Try to populate primary email/phone columns so they aren't blank
                                if email_col_right and "email" in base_df.columns:
                                    new_row["email"] = right_row.get(email_col_right)
                                if phone_col_right and "phoneNumber" in base_df.columns:
                                    new_row["phoneNumber"] = right_row.get(phone_col_right)
                                    
                                for col in cols_to_add:
                                    new_row[col_mapping[col]] = right_row[col]
                                    
                                unmatched_rows.append(new_row)
                                
                        if unmatched_rows:
                            # Add unmatched rows to the bottom
                            unmatched_df = pd.DataFrame(unmatched_rows)
                            base_df = pd.concat([base_df, unmatched_df], ignore_index=True)
                            
                        return base_df

                    # Load base file
                    if file_filtered.name.endswith(".csv"):
                        base_df = pd.read_csv(file_filtered)
                    else:
                        base_df = pd.read_excel(file_filtered)
                        
                    base_df["email_lower"] = base_df["email"].astype(str).str.lower().str.strip()
                    base_df["phone_str"] = base_df["phoneNumber"].astype(str).str.replace(r'\D', '', regex=True)

                    # --- 1. MAIN DATA ---
                    if file_main:
                        main_df = pd.read_csv(file_main) if file_main.name.endswith(".csv") else pd.read_excel(file_main)
                        cols_to_extract = ["Registered Number", "Registered mail", "batch name", "Reason"]
                        available_cols = [c for c in cols_to_extract if c in main_df.columns]
                        main_df = main_df[available_cols].copy()
                        main_df = main_df.rename(columns={"batch name": "MainData_Batch", "Reason": "MainData_Reason"})
                        
                        base_df = merge_datasets(base_df, main_df, email_col_right="Registered mail", phone_col_right="Registered Number")

                    # --- 2. LMS EXPORT BATCH DATA ---
                    if file_lms:
                        lms_df = pd.read_csv(file_lms) if file_lms.name.endswith(".csv") else pd.read_excel(file_lms)
                        cols_to_extract = ["Name", "Email", "Phone Number", "Profession", "Batches"]
                        available_cols = [c for c in cols_to_extract if c in lms_df.columns]
                        lms_df = lms_df[available_cols].copy()
                        lms_df = lms_df.rename(columns={"Name": "LMS_Name", "Profession": "LMS_Profession", "Batches": "LMS_Batches"})
                        
                        base_df = merge_datasets(base_df, lms_df, email_col_right="Email", phone_col_right="Phone Number")

                    # --- 3. AGENT BATCH CHANGE DATA ---
                    if file_agent:
                        agent_df = pd.read_csv(file_agent) if file_agent.name.endswith(".csv") else pd.read_excel(file_agent)
                        cols_to_extract = [
                            "Email Address", "Customer Email", "Customer Country Code", 
                            "Customer Contact", "Request Type", "Batch to Pause or Resume", 
                            "Customer Acknowledgement &  LMS Update Proof", 
                            "New Contact or Updated Contact of Customer", 
                            "New or Updated Country Code Of Customer", 
                            "Previous Batch", "Updated Batch"
                        ]
                        available_cols = [c for c in cols_to_extract if c in agent_df.columns]
                        agent_df = agent_df[available_cols].copy()
                        agent_df = agent_df.add_prefix("AgentChange_")
                        
                        email_col = "AgentChange_Customer Email" if "AgentChange_Customer Email" in agent_df.columns else ("AgentChange_Email Address" if "AgentChange_Email Address" in agent_df.columns else None)
                        phone_col = "AgentChange_Customer Contact" if "AgentChange_Customer Contact" in agent_df.columns else None
                        
                        base_df = merge_datasets(base_df, agent_df, email_col_right=email_col, phone_col_right=phone_col)

                    # --- 4. DATA PROCESS FILE ---
                    if file_process:
                        dp_df = pd.read_csv(file_process) if file_process.name.endswith(".csv") else pd.read_excel(file_process)
                        cols_to_extract = ["Phone", "Email", "Batch", "IntroDate"]
                        available_cols = [c for c in cols_to_extract if c in dp_df.columns]
                        dp_df = dp_df[available_cols].copy()
                        dp_df = dp_df.rename(columns={"Batch": "DataProcess_Batch", "IntroDate": "DataProcess_IntroDate"})
                        
                        base_df = merge_datasets(base_df, dp_df, email_col_right="Email", phone_col_right="Phone")
                            
                    # --- 5. BATCH PREFERENCE FILE ---
                    if file_pref:
                        pref_df = pd.read_csv(file_pref) if file_pref.name.endswith(".csv") else pd.read_excel(file_pref)
                        email_col = None
                        batch_col = None
                        phone_col = None
                        for col in pref_df.columns:
                            col_str = str(col).lower()
                            if "email" in col_str: email_col = col
                            if "preferred batch" in col_str: batch_col = col
                            if "phone" in col_str or "number" in col_str: phone_col = col
                                
                        if batch_col and (email_col or phone_col):
                            extract_cols = [batch_col]
                            if email_col: extract_cols.append(email_col)
                            if phone_col: extract_cols.append(phone_col)
                            
                            pref_extract = pref_df[list(set(extract_cols))].copy()
                            pref_extract = pref_extract.rename(columns={batch_col: "Preferred_Batch"})
                            
                            base_df = merge_datasets(base_df, pref_extract, email_col_right=email_col, phone_col_right=phone_col)

                    # Cleanup helper columns
                    base_df = base_df.drop(columns=["email_lower", "phone_str"], errors="ignore")
                    
                    st.success("Files merged successfully!")
                    st.dataframe(base_df.head(15))
                    
                    # Download button
                    output2 = io.BytesIO()
                    with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                        base_df.to_excel(writer, index=False, sheet_name='Merged_Data')
                    final_data = output2.getvalue()
                    
                    st.download_button(
                        label="Download Final Merged Data as Excel",
                        data=final_data,
                        file_name="Final_Merged_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    import traceback
                    st.error(f"Error merging files: {e}")
                    st.text(traceback.format_exc())
