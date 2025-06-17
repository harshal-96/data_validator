import streamlit as st
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
from io import BytesIO
import io

st.set_page_config(
    page_title="Excel Data Processor",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Data Processing & Validation App")
st.markdown("Upload two Excel files to merge and validate loan data")

# File upload section
col1, col2 = st.columns(2)

with col1:
    st.subheader("Upload Asset Details Excel")
    uploaded_file1 = st.file_uploader(
        "Choose Asset Details Excel file (Should contain sheets named 'Aadhar', 'pancard', and 'Asset')", 
        type=['xlsx', 'xls'],
        key="file1"
    )

with col2:
    st.subheader("Upload Asset Data Excel")
    uploaded_file2 = st.file_uploader(
        "Choose Asset Data Excel file (Should contain sheets named 'ApplicationForm', 'Applicant', and 'Asset')", 
        type=['xlsx', 'xls'],
        key="file2"
    )

def convert_date(val):
    """Convert date format"""
    val = str(val).zfill(8)
    return f"{val[:2]}-{val[2:4]}-{val[4:]}"

def normalize_name(name):
    """Normalize name for comparison"""
    return ' '.join(str(name).upper().split())

def check_name_similarity(app_names, owner_names, threshold=85):
    """Check name similarity using fuzzy matching"""
    matched = []
    for app_name in app_names:
        for owner_name in owner_names:
            score = fuzz.token_set_ratio(normalize_name(app_name), normalize_name(owner_name))
            if score >= threshold:
                matched.append((app_name, owner_name, score))
    return matched

def process_data(file1, file2):
    """Main data processing function"""
    try:
        # Read first Excel file (Asset Details)
        excel_file1 = pd.ExcelFile(file1)
        
        # Parse sheets from first file
        aadhar_df = excel_file1.parse('Aadhar') if 'Aadhar' in excel_file1.sheet_names else pd.DataFrame()
        pancard_df = excel_file1.parse('pancard') if 'pancard' in excel_file1.sheet_names else pd.DataFrame()
        asset_df = excel_file1.parse('Asset') if 'Asset' in excel_file1.sheet_names else pd.DataFrame()
        
        # Read second Excel file (Asset Data)
        excel_file2 = pd.ExcelFile(file2)
        
        # Parse sheets from second file
        application_form_df = excel_file2.parse('ApplicationForm') if 'ApplicationForm' in excel_file2.sheet_names else pd.DataFrame()
        applicant_details_df = excel_file2.parse('Applicant') if 'Applicant' in excel_file2.sheet_names else pd.DataFrame()
        asset_details_df = excel_file2.parse('Asset') if 'Asset' in excel_file2.sheet_names else pd.DataFrame()
        
        st.success("✅ Files loaded successfully!")
        
        # Show available sheets
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Asset Details File Sheets:**", excel_file1.sheet_names)
        with col2:
            st.write("**Asset Data File Sheets:**", excel_file2.sheet_names)
        
        # Data cleaning and preprocessing
        if not aadhar_df.empty:
            aadhar_df = aadhar_df.drop(columns=['aadhaar_number'], errors='ignore')
            
        if not pancard_df.empty:
            pancard_df = pancard_df.drop(columns=['full_name_split', 'masked_aadhaar', 'pan_number', 'pan_number.1', 'phone_number'], errors='ignore')
        
        # Standardize column names
        column_mappings = {
            'PartnerLoanNumber': 'Loan Number',
            'AadhaarNumber': 'Aadhar Number',
            'PancardNumber': 'Pancard Number',
            'dob': 'DOB'
        }
        
        for df in [aadhar_df, pancard_df, asset_df]:
            if not df.empty:
                df.rename(columns=column_mappings, inplace=True)
        
        # Process applicant details if available
        if not applicant_details_df.empty:
            # Convert DOB format
            if 'DOB' in applicant_details_df.columns:
                applicant_details_df['DOB'] = applicant_details_df['DOB'].apply(convert_date)
                applicant_details_df['DOB'] = pd.to_datetime(applicant_details_df['DOB'], format='%d-%m-%Y', errors='coerce')
            
            # Create full name
            if all(col in applicant_details_df.columns for col in ['First Name', 'Middle Name', 'Last Name']):
                applicant_details_df['Full Name'] = (
                    applicant_details_df['First Name'].fillna('').astype(str) + ' ' +
                    applicant_details_df['Middle Name'].fillna('').astype(str) + ' ' +
                    applicant_details_df['Last Name'].fillna('').astype(str)
                ).str.strip()
            
            # Normalize Pancard Number
            if 'Pancard Number' in applicant_details_df.columns:
                applicant_details_df['Pancard Number'] = applicant_details_df['Pancard Number'].str.upper()
        
        # Process asset dataframe
        if not asset_df.empty and all(col in asset_df.columns for col in ['owner_name', 'father_name']):
            asset_df['full_name'] = (
                asset_df['owner_name'].fillna('').astype(str) + ' ' +
                asset_df['father_name'].fillna('').astype(str)
            ).str.strip()
        
        # Remove duplicates
        if not pancard_df.empty:
            pancard_df = pancard_df.drop_duplicates()
        
        # Create borrowers dataframe
        borrowers_df = pd.DataFrame()
        if not applicant_details_df.empty and 'Applicant category(borrower/co-borrower/guarantor)' in applicant_details_df.columns:
            borrowers_df = applicant_details_df[
                applicant_details_df['Applicant category(borrower/co-borrower/guarantor)'].str.lower() == 'borrower'
            ]
        
        return {
            'aadhar_df': aadhar_df,
            'pancard_df': pancard_df,
            'asset_df': asset_df,
            'application_form_df': application_form_df,
            'applicant_details_df': applicant_details_df,
            'asset_details_df': asset_details_df,
            'borrowers_df': borrowers_df
        }
        
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None

def validate_data(dataframes):
    """Perform data validation checks"""
    validation_results = []
    
    aadhar_df = dataframes['aadhar_df']
    pancard_df = dataframes['pancard_df']
    asset_df = dataframes['asset_df']
    applicant_details_df = dataframes['applicant_details_df']
    
    # Validation 1: Aadhaar Number consistency
    if not applicant_details_df.empty and not aadhar_df.empty:
        if 'Loan Number' in applicant_details_df.columns and 'Aadhar Number' in applicant_details_df.columns:
            app_map = applicant_details_df.groupby('Loan Number')['Aadhar Number'].apply(set)
            aadhaar_map = aadhar_df.groupby('Loan Number')['Aadhar Number'].apply(set)
            common_loans = set(app_map.index).intersection(set(aadhaar_map.index))
            
            aadhaar_mismatches = []
            for loan in common_loans:
                app_set = app_map[loan]
                aadhaar_set = aadhaar_map[loan]
                if app_set != aadhaar_set:
                    combined_set = app_set.union(aadhaar_set)
                    if len(combined_set) > 2:
                        aadhaar_mismatches.append({
                            'loan': loan,
                            'applicant_aadhaar': list(app_set),
                            'aadhaar_file': list(aadhaar_set),
                            'total_unique': len(combined_set)
                        })
            
            validation_results.append({
                'check': 'Aadhaar Number Consistency',
                'status': 'Pass' if len(aadhaar_mismatches) == 0 else 'Fail',
                'details': aadhaar_mismatches,
                'count': len(aadhaar_mismatches)
            })
    
    # Validation 2: PAN Number consistency
    if not applicant_details_df.empty and not pancard_df.empty:
        if all(col in applicant_details_df.columns for col in ['Loan Number', 'Pancard Number']) and 'Pancard Number' in pancard_df.columns:
            app_map = applicant_details_df.groupby('Loan Number')['Pancard Number'].apply(set)
            pancard_map = pancard_df.groupby('Loan Number')['Pancard Number'].apply(set)
            common_loans = set(app_map.index).intersection(set(pancard_map.index))
            
            pan_mismatches = []
            for loan in common_loans:
                app_set = app_map[loan]
                pancard_set = pancard_map[loan]
                if app_set != pancard_set:
                    combined_set = app_set.union(pancard_set)
                    if len(combined_set) > 2:
                        pan_mismatches.append({
                            'loan': loan,
                            'applicant_pan': list(app_set),
                            'pancard_file': list(pancard_set),
                            'total_unique': len(combined_set)
                        })
            
            validation_results.append({
                'check': 'PAN Number Consistency',
                'status': 'Pass' if len(pan_mismatches) == 0 else 'Fail',
                'details': pan_mismatches,
                'count': len(pan_mismatches)
            })
    
    # Validation 3: Name consistency (Applicant vs Asset Owner)
    if not applicant_details_df.empty and not asset_df.empty:
        if all(col in applicant_details_df.columns for col in ['Loan Number', 'Full Name']) and all(col in asset_df.columns for col in ['Loan Number', 'full_name']):
            app_name_map = applicant_details_df.groupby('Loan Number')['Full Name'].apply(set)
            owner_name_map = asset_df.groupby('Loan Number')['full_name'].apply(set)
            common_loans = set(app_name_map.index).intersection(set(owner_name_map.index))
            
            name_mismatches = []
            for loan in common_loans:
                app_names = app_name_map[loan]
                owner_names = owner_name_map[loan]
                combined_names = app_names.union(owner_names)
                
                if app_names != owner_names and len(combined_names) > 2:
                    # Check for fuzzy matches
                    matches = check_name_similarity(app_names, owner_names)
                    if not matches:  # Only report if no fuzzy matches found
                        name_mismatches.append({
                            'loan': loan,
                            'applicant_names': list(app_names),
                            'asset_owner_names': list(owner_names),
                            'total_unique': len(combined_names)
                        })
            
            validation_results.append({
                'check': 'Name Consistency (Applicant vs Asset Owner)',
                'status': 'Pass' if len(name_mismatches) == 0 else 'Fail',
                'details': name_mismatches,
                'count': len(name_mismatches)
            })
    
    # Validation 4: DOB consistency (Applicant vs PAN)
    if not applicant_details_df.empty and not pancard_df.empty:
        if all(col in applicant_details_df.columns for col in ['Loan Number', 'DOB']) and all(col in pancard_df.columns for col in ['Loan Number', 'DOB']):
            app_dob_map = applicant_details_df.groupby('Loan Number')['DOB'].apply(set)
            pancard_dob_map = pancard_df.groupby('Loan Number')['DOB'].apply(set)
            common_loans = set(app_dob_map.index).intersection(set(pancard_dob_map.index))
            
            dob_mismatches = []
            for loan in common_loans:
                app_dobs = app_dob_map[loan]
                pancard_dobs = pancard_dob_map[loan]
                
                if app_dobs != pancard_dobs:
                    combined_dobs = app_dobs.union(pancard_dobs)
                    if len(combined_dobs) > 2:
                        dob_mismatches.append({
                            'loan': loan,
                            'applicant_dobs': [str(d) for d in app_dobs],
                            'pancard_dobs': [str(d) for d in pancard_dobs],
                            'total_unique': len(combined_dobs)
                        })
            
            validation_results.append({
                'check': 'DOB Consistency (Applicant vs PAN)',
                'status': 'Pass' if len(dob_mismatches) == 0 else 'Fail',
                'details': dob_mismatches,
                'count': len(dob_mismatches)
            })
    
    # Validation 5: Mobile Number consistency (Applicant vs Asset)
    if not applicant_details_df.empty and not asset_df.empty:
        if all(col in applicant_details_df.columns for col in ['Loan Number', 'Mobile Number']) and all(col in asset_df.columns for col in ['Loan Number', 'mobile_number']):
            app_mobile_map = applicant_details_df.groupby('Loan Number')['Mobile Number'].apply(set)
            asset_mobile_map = asset_df.groupby('Loan Number')['mobile_number'].apply(set)
            common_loans = set(app_mobile_map.index).intersection(set(asset_mobile_map.index))
            
            mobile_mismatches = []
            for loan in common_loans:
                app_mobiles = app_mobile_map[loan]
                asset_mobiles = asset_mobile_map[loan]
                
                if app_mobiles != asset_mobiles:
                    combined_mobiles = app_mobiles.union(asset_mobiles)
                    if len(combined_mobiles) > 2:
                        mobile_mismatches.append({
                            'loan': loan,
                            'applicant_mobiles': list(app_mobiles),
                            'asset_mobiles': list(asset_mobiles),
                            'total_unique': len(combined_mobiles)
                        })
            
            validation_results.append({
                'check': 'Mobile Number Consistency (Applicant vs Asset)',
                'status': 'Pass' if len(mobile_mismatches) == 0 else 'Fail',
                'details': mobile_mismatches,
                'count': len(mobile_mismatches)
            })
    
    return validation_results

def create_final_dataframe(dataframes):
    """Create the final merged dataframe"""
    try:
        aadhar_df = dataframes['aadhar_df']
        pancard_df = dataframes['pancard_df']
        asset_df = dataframes['asset_df']
        applicant_details_df = dataframes['applicant_details_df']
        application_form_df = dataframes['application_form_df']
        borrowers_df = dataframes['borrowers_df']
        asset_df=asset_df.drop_duplicates(subset=['Loan Number'], keep='first')
        if borrowers_df.empty or asset_df.empty:
            return pd.DataFrame()
        
        # Create borrower identity dataframe
        borrower_aadhar = pd.DataFrame()
        borrower_pan = pd.DataFrame()
        
        if not aadhar_df.empty and 'Aadhar Number' in borrowers_df.columns:
            borrower_aadhar = pd.merge(
                borrowers_df[['Loan Number', 'Aadhar Number']],
                aadhar_df[['Loan Number', 'Aadhar Number', 'age_range', 'state']],
                on=['Loan Number', 'Aadhar Number'],
                how='left'
            )
        
        if not pancard_df.empty and all(col in borrowers_df.columns for col in ['Pancard Number', 'DOB']):
            borrower_pan = pd.merge(
                borrowers_df[['Loan Number', 'Pancard Number', 'DOB']],
                pancard_df[['Loan Number', 'Pancard Number', 'full_name', 'DOB']],
                on=['Loan Number', 'Pancard Number'],
                how='left'
            )
        
        # Merge borrower identity data
        if not borrower_aadhar.empty and not borrower_pan.empty:
            borrower_identity_df = pd.merge(borrower_aadhar, borrower_pan, on='Loan Number', how='outer')
        elif not borrower_aadhar.empty:
            borrower_identity_df = borrower_aadhar
        elif not borrower_pan.empty:
            borrower_identity_df = borrower_pan
        else:
            borrower_identity_df = borrowers_df[['Loan Number']]
        
        # Trim asset dataframe
        asset_columns = ['Loan Number', 'rc_number', 'registration_date', 'owner_name', 'mobile_number',
                        'vehicle_category', 'vehicle_chasi_number', 'maker_description', 'maker_model',
                        'color', 'fuel_type', 'manufacturing_date_formatted', 'insurance_company',
                        'insurance_upto', 'permit_number', 'blacklist_status', 'rc_status', 'rto_code']
        
        available_columns = [col for col in asset_columns if col in asset_df.columns]
        asset_df_trimmed = asset_df[available_columns]
        
        # Merge asset data with borrower identity
        final_df = pd.merge(asset_df_trimmed, borrower_identity_df, on='Loan Number', how='left')
        
        # Add missing loan numbers from application form
        if not application_form_df.empty:
            missing_loan_numbers = application_form_df[~application_form_df['Loan Number'].isin(final_df['Loan Number'])]
            if not missing_loan_numbers.empty:
                missing_rows = pd.DataFrame(columns=final_df.columns)
                missing_rows['Loan Number'] = missing_loan_numbers['Loan Number'].values
                final_df = pd.concat([final_df, missing_rows], ignore_index=True)
        
        return final_df.reset_index(drop=True)
        
    except Exception as e:
        st.error(f"Error creating final dataframe: {str(e)}")
        return pd.DataFrame()

# Main processing
if uploaded_file1 is not None and uploaded_file2 is not None:
    with st.spinner("Processing files..."):
        dataframes = process_data(uploaded_file1, uploaded_file2)
        
        if dataframes:
            st.success("Files processed successfully!")
            
            # Show data preview
            st.subheader("📋 Data Preview")
            tab1, tab2, tab3, tab4 = st.tabs(["Applicant Details", "Asset Data", "Aadhaar Data", "PAN Data"])
            
            with tab1:
                if not dataframes['applicant_details_df'].empty:
                    st.write(f"Shape: {dataframes['applicant_details_df'].shape}")
                    st.dataframe(dataframes['applicant_details_df'].head())
                else:
                    st.info("No applicant details data available")
            
            with tab2:
                if not dataframes['asset_df'].empty:
                    st.write(f"Shape: {dataframes['asset_df'].shape}")
                    st.dataframe(dataframes['asset_df'].head())
                else:
                    st.info("No asset data available")
            
            with tab3:
                if not dataframes['aadhar_df'].empty:
                    st.write(f"Shape: {dataframes['aadhar_df'].shape}")
                    st.dataframe(dataframes['aadhar_df'].head())
                else:
                    st.info("No Aadhaar data available")
            
            with tab4:
                if not dataframes['pancard_df'].empty:
                    st.write(f"Shape: {dataframes['pancard_df'].shape}")
                    st.dataframe(dataframes['pancard_df'].head())
                else:
                    st.info("No PAN data available")
            
            # Data validation
            st.subheader("🔍 Data Validation Results")
            validation_results = validate_data(dataframes)
            
            # Summary of validation results
            total_checks = len(validation_results)
            passed_checks = sum(1 for result in validation_results if result['status'] == 'Pass')
            failed_checks = total_checks - passed_checks
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Checks", total_checks)
            with col2:
                st.metric("Passed", passed_checks, delta=None, delta_color="normal")
            with col3:
                st.metric("Failed", failed_checks, delta=None, delta_color="inverse")
            
            # Detailed validation results
            for result in validation_results:
                if result['status'] == 'Pass':
                    st.success(f"✅ {result['check']}: {result['status']}")
                else:
                    st.error(f"❌ {result['check']}: {result['status']} ({result['count']} mismatches found)")
                    
                    if result['details']:
                        with st.expander(f"View {result['check']} Mismatches ({result['count']} issues)"):
                            for i, detail in enumerate(result['details'], 1):
                                st.write(f"**Mismatch {i}:**")
                                st.write(f"- **Loan Number:** {detail['loan']}")
                                
                                if 'applicant_aadhaar' in detail:
                                    st.write(f"- **Applicant Aadhaar:** {detail['applicant_aadhaar']}")
                                    st.write(f"- **Aadhaar File:** {detail['aadhaar_file']}")
                                
                                elif 'applicant_pan' in detail:
                                    st.write(f"- **Applicant PAN:** {detail['applicant_pan']}")
                                    st.write(f"- **PAN File:** {detail['pancard_file']}")
                                
                                elif 'applicant_names' in detail:
                                    st.write(f"- **Applicant Names:** {detail['applicant_names']}")
                                    st.write(f"- **Asset Owner Names:** {detail['asset_owner_names']}")
                                
                                elif 'applicant_dobs' in detail:
                                    st.write(f"- **Applicant DOBs:** {detail['applicant_dobs']}")
                                    st.write(f"- **PAN File DOBs:** {detail['pancard_dobs']}")
                                
                                elif 'applicant_mobiles' in detail:
                                    st.write(f"- **Applicant Mobiles:** {detail['applicant_mobiles']}")
                                    st.write(f"- **Asset File Mobiles:** {detail['asset_mobiles']}")
                                
                                st.write(f"- **Total Unique Values:** {detail['total_unique']}")
                                st.write("---")
            
            # Show validation summary table
            if validation_results:
                st.subheader("📊 Validation Summary Table")
                summary_data = []
                for result in validation_results:
                    summary_data.append({
                        'Validation Check': result['check'],
                        'Status': result['status'],
                        'Issues Found': result.get('count', 0),
                        'Details Available': 'Yes' if result['details'] else 'No'
                    })
                
                summary_df = pd.DataFrame(summary_data)
                st.dataframe(summary_df, use_container_width=True)
            
            # Create final dataframe
            st.subheader("📊 Final Merged Dataset")
            final_df = create_final_dataframe(dataframes)
            
            if not final_df.empty:
                st.write(f"Final dataset shape: {final_df.shape}")
                st.dataframe(final_df)
                
                # Download option
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Sheet1')
                output.seek(0)  # Reset buffer position to the beginning

                # Download button
                st.download_button(
                    label="📥 Download Final Dataset as Excel",
                    data=output,
                    file_name="merged_loan_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Summary statistics
                st.subheader("📈 Summary Statistics")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total Records", final_df.shape[0])
                
                with col2:
                    if 'Loan Number' in final_df.columns:
                        st.metric("Unique Loan Numbers", final_df['Loan Number'].nunique())
                
                with col3:
                    missing_percentage = (final_df.isnull().sum().sum() / (final_df.shape[0] * final_df.shape[1])) * 100
                    st.metric("Missing Data %", f"{missing_percentage:.1f}%")
            
            else:
                st.warning("Could not create final merged dataset. Please check your data structure.")

else:
    st.info("👆 Please upload both Excel files to begin processing")
    
    # Show instructions
    with st.expander("📝 Instructions"):
        st.markdown("""
        ### How to use this app:
        
        1. **Upload Asset Details Excel**: Should contain sheets named 'Aadhar', 'pancard', and 'Asset'
        2. **Upload Asset Data Excel**: Should contain sheets named 'ApplicationForm', 'Applicant', and 'Asset'
        
        ### Expected sheet structure:
        
        **Asset Details Excel:**
        - **Aadhar sheet**: Contains Aadhaar information with columns like 'PartnerLoanNumber', 'AadhaarNumber', 'age_range', 'state'
        - **pancard sheet**: Contains PAN information with columns like 'PartnerLoanNumber', 'PancardNumber', 'dob', 'full_name'
        - **Asset sheet**: Contains asset information with columns like 'PartnerLoanNumber', 'owner_name', 'father_name', etc.
        
        **Asset Data Excel:**
        - **ApplicationForm sheet**: Contains loan application data
        - **Applicant sheet**: Contains applicant details with columns like 'Loan Number', 'First Name', 'Last Name', 'Aadhar Number', 'Pancard Number', 'DOB'
        - **Asset sheet**: Contains additional asset details
        
        ### Features:
        - ✅ Data validation and consistency checks
        - 📊 Interactive data preview
        - 🔄 Automatic data merging and standardization
        - 📥 Download processed results as Excel
        - 📈 Summary statistics and insights
        """)