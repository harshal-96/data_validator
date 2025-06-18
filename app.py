import streamlit as st
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
import io
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Excel Sheet Validation Tool",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Sheet Validation Tool")
st.markdown("Upload two Excel files to validate data consistency across different sheets")

# Sidebar for file uploads
st.sidebar.header("File Upload")
api_file = st.sidebar.file_uploader("Upload API Excel File (AssetDetails)", type=['xlsx'], key="api_file")
manual_file = st.sidebar.file_uploader("Upload Manual Excel File (Asset_a978...)", type=['xlsx'], key="manual_file")

# Utility functions
def normalize(name):
    """Normalize name for comparison"""
    if pd.isna(name):
        return ""
    return ' '.join(str(name).upper().split())

def convert_date(val):
    """Convert date format"""
    if pd.isna(val):
        return val
    val = str(val).zfill(8)
    return f"{val[:2]}-{val[2:4]}-{val[4:]}"

def check_name_similarity(app_names, owner_names, threshold=85):
    """Check name similarity using fuzzy matching"""
    matched = []
    for app_name in app_names:
        for owner_name in owner_names:
            score = fuzz.token_set_ratio(normalize(app_name), normalize(owner_name))
            if score >= threshold:
                matched.append((app_name, owner_name, score))
    return matched

def load_and_process_data(api_file, manual_file):
    """Load and process data from both Excel files"""
    try:
        # Load API file sheets
        api_excel = pd.ExcelFile(api_file)
        aadhar_df = api_excel.parse('Aadhar')
        pancard_df = api_excel.parse('pancard')
        asset_df = api_excel.parse('Asset')
        
        # Load Manual file sheets
        manual_excel = pd.ExcelFile(manual_file)
        application_form_df = manual_excel.parse('ApplicationForm')
        applicant_details_df = manual_excel.parse('Applicant')
        asset_details_df = manual_excel.parse('Asset')
        
        # Data preprocessing
        # Drop sensitive columns
        aadhar_df = aadhar_df.drop(columns=['aadhaar_number'], errors='ignore')
        pancard_df = pancard_df.drop(columns=['full_name_split', 'masked_aadhaar', 'pan_number', 'pan_number.1', 'phone_number'], errors='ignore')
        
        # Standardize column names
        asset_df = asset_df.rename(columns={'PartnerLoanNumber': 'Loan Number'})
        pancard_df = pancard_df.rename(columns={'PartnerLoanNumber': 'Loan Number'})
        aadhar_df = aadhar_df.rename(columns={'PartnerLoanNumber': 'Loan Number'})
        
        # Create full names
        applicant_details_df['Full Name'] = (
            applicant_details_df['First Name'].fillna('').astype(str) + ' ' +
            applicant_details_df['Middle Name'].fillna('').astype(str) + ' ' +
            applicant_details_df['Last Name'].fillna('').astype(str)
        ).str.strip()
        
        asset_df['full_name'] = (
            asset_df['owner_name'].fillna('').astype(str) + ' ' +
            asset_df['father_name'].fillna('').astype(str)
        ).str.strip()
        
        # Convert dates
        applicant_details_df['DOB'] = applicant_details_df['DOB'].apply(convert_date)
        applicant_details_df['DOB'] = pd.to_datetime(applicant_details_df['DOB'], format='%d-%m-%Y', errors='coerce')
        
        # Standardize PAN numbers
        applicant_details_df['Pancard Number'] = applicant_details_df['Pancard Number'].str.upper()
        
        # Remove duplicates
        pancard_df = pancard_df.drop_duplicates()
        
        # Rename columns for consistency
        aadhar_df = aadhar_df.rename(columns={'AadhaarNumber': 'Aadhar Number'})
        pancard_df = pancard_df.rename(columns={'PancardNumber': 'Pancard Number', 'dob': 'DOB'})
        
        return {
            'aadhar_df': aadhar_df,
            'pancard_df': pancard_df,
            'asset_df': asset_df,
            'application_form_df': application_form_df,
            'applicant_details_df': applicant_details_df,
            'asset_details_df': asset_details_df
        }
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def validate_aadhar_section(aadhar_df, applicant_details_df):
    """Validate Aadhar section"""
    st.subheader("🆔 Aadhar Validation")
    
    mismatches = []
    
    # Create mapping from Loan Number → Aadhaar Numbers
    app_map = applicant_details_df.groupby('Loan Number')['Aadhar Number'].apply(set)
    aadhaar_map = aadhar_df.groupby('Loan Number')['AadhaarNumber'].apply(set) if 'AadhaarNumber' in aadhar_df.columns else aadhar_df.groupby('Loan Number')['Aadhar Number'].apply(set)
    
    # Get common loan numbers
    common_loans = set(app_map.index).intersection(set(aadhaar_map.index))
    
    for loan in common_loans:
        app_set = app_map[loan]
        aadhaar_set = aadhaar_map[loan]
        
        if app_set != aadhaar_set:
            combined_set = app_set.union(aadhaar_set)
            if len(combined_set) > 2:
                mismatches.append({
                    'Loan Number': loan,
                    'Issue': 'Aadhar Number Mismatch',
                    'Applicant Details': list(app_set),
                    'Aadhar Sheet': list(aadhaar_set),
                    'Total Unique': len(combined_set)
                })
    
    if mismatches:
        st.error(f"Found {len(mismatches)} Aadhar mismatches")
        mismatch_df = pd.DataFrame(mismatches)
        st.dataframe(mismatch_df, use_container_width=True)
    else:
        st.success("✅ No Aadhar mismatches found")
    
    return mismatches

def validate_pancard_section(pancard_df, applicant_details_df):
    """Validate PAN Card section"""
    st.subheader("💳 PAN Card Validation")
    
    mismatches = []
    
    # Create mapping from Loan Number → Pancard Numbers
    app_map = applicant_details_df.groupby('Loan Number')['Pancard Number'].apply(set)
    pancard_map = pancard_df.groupby('Loan Number')['Pancard Number'].apply(set) if 'Pancard Number' in pancard_df.columns else pancard_df.groupby('Loan Number')['PancardNumber'].apply(set)
    
    # Get common loan numbers
    common_loans = set(app_map.index).intersection(set(pancard_map.index))
    
    for loan in common_loans:
        app_set = app_map[loan]
        pancard_set = pancard_map[loan]
        
        if app_set != pancard_set:
            combined_set = app_set.union(pancard_set)
            if len(combined_set) > 2:
                mismatches.append({
                    'Loan Number': loan,
                    'Issue': 'PAN Card Number Mismatch',
                    'Applicant Details': list(app_set),
                    'PAN Card Sheet': list(pancard_set),
                    'Total Unique': len(combined_set)
                })
    
    if mismatches:
        st.error(f"Found {len(mismatches)} PAN Card mismatches")
        mismatch_df = pd.DataFrame(mismatches)
        st.dataframe(mismatch_df, use_container_width=True)
    else:
        st.success("✅ No PAN Card mismatches found")
    
    return mismatches

def validate_name_section(asset_df, applicant_details_df):
    """Validate Name section"""
    st.subheader("👤 Name Validation")
    
    mismatches = []
    
    app_name_map = applicant_details_df.groupby('Loan Number')['Full Name'].apply(set)
    owner_name_map = asset_df.groupby('Loan Number')['full_name'].apply(set)
    
    # Get common loan numbers
    common_loans = set(app_name_map.index).intersection(set(owner_name_map.index))
    
    for loan in common_loans:
        app_names = app_name_map[loan]
        owner_names = owner_name_map[loan]
        combined_names = app_names.union(owner_names)
        
        if app_names != owner_names and len(combined_names) > 2:
            matches = check_name_similarity(app_names, owner_names)
            if not matches:
                mismatches.append({
                    'Loan Number': loan,
                    'Issue': 'Name Mismatch',
                    'Applicant Names': list(app_names),
                    'Asset Owner Names': list(owner_names)
                })
    
    if mismatches:
        st.error(f"Found {len(mismatches)} Name mismatches")
        mismatch_df = pd.DataFrame(mismatches)
        st.dataframe(mismatch_df, use_container_width=True)
    else:
        st.success("✅ No Name mismatches found")
    
    return mismatches

def validate_dob_section(pancard_df, applicant_details_df):
    """Validate DOB section"""
    st.subheader("📅 Date of Birth Validation")
    
    mismatches = []
    
    # Create mapping from Loan Number → DOBs
    app_dob_map = applicant_details_df.groupby('Loan Number')['DOB'].apply(set)
    pancard_dob_map = pancard_df.groupby('Loan Number')['DOB'].apply(set) if 'DOB' in pancard_df.columns else pancard_df.groupby('Loan Number')['dob'].apply(set)
    
    # Get common loan numbers
    common_loans = set(app_dob_map.index).intersection(set(pancard_dob_map.index))
    
    for loan in common_loans:
        app_dobs = app_dob_map[loan]
        pancard_dobs = pancard_dob_map[loan]
        
        if app_dobs != pancard_dobs:
            combined_dobs = app_dobs.union(pancard_dobs)
            if len(combined_dobs) > 2:
                mismatches.append({
                    'Loan Number': loan,
                    'Issue': 'DOB Mismatch',
                    'Applicant DOB': list(app_dobs),
                    'PAN Card DOB': list(pancard_dobs),
                    'Total Unique': len(combined_dobs)
                })
    
    if mismatches:
        st.error(f"Found {len(mismatches)} DOB mismatches")
        mismatch_df = pd.DataFrame(mismatches)
        st.dataframe(mismatch_df, use_container_width=True)
    else:
        st.success("✅ No DOB mismatches found")
    
    return mismatches

def validate_mobile_section(asset_df, applicant_details_df):
    """Validate Mobile Number section"""
    st.subheader("📱 Mobile Number Validation")
    
    mismatches = []
    
    # Group by Loan Number → Mobile Numbers
    app_mobile_map = applicant_details_df.groupby('Loan Number')['Mobile Number'].apply(set)
    asset_mobile_map = asset_df.groupby('Loan Number')['mobile_number'].apply(set)
    
    # Get common loan numbers
    common_loans = set(app_mobile_map.index).intersection(set(asset_mobile_map.index))
    
    for loan in common_loans:
        app_mobiles = app_mobile_map[loan]
        asset_mobiles = asset_mobile_map[loan]
        
        if app_mobiles != asset_mobiles:
            combined_mobiles = app_mobiles.union(asset_mobiles)
            if len(combined_mobiles) > 2:
                mismatches.append({
                    'Loan Number': loan,
                    'Issue': 'Mobile Number Mismatch',
                    'Applicant Mobile': list(app_mobiles),
                    'Asset Mobile': list(asset_mobiles),
                    'Total Unique': len(combined_mobiles)
                })
    
    if mismatches:
        st.error(f"Found {len(mismatches)} Mobile Number mismatches")
        mismatch_df = pd.DataFrame(mismatches)
        st.dataframe(mismatch_df, use_container_width=True)
    else:
        st.success("✅ No Mobile Number mismatches found")
    
    return mismatches

def validate_application_form_section(asset_df, application_form_df):
    """Validate Application Form section - Name and Mobile Number"""
    
    # Name Validation
    st.subheader("👤 Name Validation (Application Form)")
    name_mismatches = []
    
    
    mobile_mismatches = []
    
    # Get common loan numbers
    common_loans = set(asset_df['Loan Number']).intersection(set(application_form_df['Loan Number']))
    
    for loan in common_loans:
        asset_row = asset_df[asset_df['Loan Number'] == loan].iloc[0]
        application_row = application_form_df[application_form_df['Loan Number'] == loan].iloc[0]
        
        # Mobile Number check
        if normalize(str(application_row.get('Mobile No.', ''))) != normalize(str(asset_row.get('mobile_number', ''))):
            mobile_mismatches.append({
                'Loan Number': loan,
                'Issue': 'Mobile Number Mismatch',
                'Application Form Mobile': str(application_row.get('Mobile No.', '')),
                'Asset Mobile': str(asset_row.get('mobile_number', ''))
            })
        
        # Customer Name check
        score_name = fuzz.token_set_ratio(
            normalize(str(application_row.get('Customer Name', ''))), 
            normalize(str(asset_row.get('full_name', '')))
        )
        if score_name < 85:
            name_mismatches.append({
                'Loan Number': loan,
                'Issue': 'Name Mismatch',
                'Application Form Name': str(application_row.get('Customer Name', '')),
                'Asset Owner Name': str(asset_row.get('full_name', '')),
                'Similarity Score': score_name
            })
    
    # Display Name validation results
    if name_mismatches:
        st.error(f"Found {len(name_mismatches)} Name mismatches")
        name_df = pd.DataFrame(name_mismatches)
        st.dataframe(name_df, use_container_width=True)
    else:
        st.success("✅ No Name mismatches found")
    
    st.markdown("---")
    # Mobile Number Validation  
    st.subheader("📱 Mobile Number Validation (Application Form)")
    # Display Mobile validation results
    if mobile_mismatches:
        st.error(f"Found {len(mobile_mismatches)} Mobile Number mismatches")
        mobile_df = pd.DataFrame(mobile_mismatches)
        st.dataframe(mobile_df, use_container_width=True)
    else:
        st.success("✅ No Mobile Number mismatches found")
    
    return name_mismatches + mobile_mismatches

def validate_asset_form_section(asset_df, asset_details_df, applicant_details_df, application_form_df):
    """Validate Asset Form section - Registration Number, RC Date, Engine No, Chassis No, Address"""
    
    # Registration Number Validation
    st.subheader("📋 Registration Number Validation")
    reg_mismatches = []
    
    date_mismatches = []
    
    engine_mismatches = []
    
    chassis_mismatches = []

    address_mismatches = []
    
    # Get common loan numbers
    common_loans = set(asset_df['Loan Number']).intersection(set(asset_details_df['Loan Number']))
    
    for loan in common_loans:
        asset_row = asset_df[asset_df['Loan Number'] == loan].iloc[0]
        details_row = asset_details_df[asset_details_df['Loan Number'] == loan].iloc[0]
        
        # Get corresponding applicant row for address validation
        applicant_row = applicant_details_df[applicant_details_df['Loan Number'] == loan].iloc[0] if loan in applicant_details_df['Loan Number'].values else None
        
        # Registration Number (RC Number)
        if normalize(str(asset_row.get('rc_number', ''))) != normalize(str(details_row.get('Registration No', ''))):
            reg_mismatches.append({
                'Loan Number': loan,
                'Issue': 'Registration Number Mismatch',
                'Asset RC Number': str(asset_row.get('rc_number', '')),
                'Asset Details Reg No': str(details_row.get('Registration No', ''))
            })
        
        # Registration Date
        if str(asset_row.get('registration_date', '')) != str(details_row.get('Registration Date', '')):
            date_mismatches.append({
                'Loan Number': loan,
                'Issue': 'Registration Date Mismatch',
                'Asset Registration Date': str(asset_row.get('registration_date', '')),
                'Asset Details Reg Date': str(details_row.get('Registration Date', ''))
            })
        
        # Engine Number
        if normalize(str(asset_row.get('vehicle_engine_number', ''))) != normalize(str(details_row.get('Engine No', ''))):
            engine_mismatches.append({
                'Loan Number': loan,
                'Issue': 'Engine Number Mismatch',
                'Asset Engine Number': str(asset_row.get('vehicle_engine_number', '')),
                'Asset Details Engine No': str(details_row.get('Engine No', ''))
            })
        
        # Chassis Number
        if normalize(str(asset_row.get('vehicle_chasi_number', ''))) != normalize(str(details_row.get('Chassis No', ''))):
            chassis_mismatches.append({
                'Loan Number': loan,
                'Issue': 'Chassis Number Mismatch',
                'Asset Chassis Number': str(asset_row.get('vehicle_chasi_number', '')),
                'Asset Details Chassis No': str(details_row.get('Chassis No', ''))
            })
        
        # Address validation (if applicant row exists)
        if applicant_row is not None:
            score = fuzz.token_set_ratio(
                normalize(str(asset_row.get('permanent_address', ''))), 
                normalize(str(applicant_row.get('Customer Address', '')))
            )
            if score < 85:
                address_mismatches.append({
                    'Loan Number': loan,
                    'Issue': 'Address Mismatch',
                    'Asset Address': str(asset_row.get('permanent_address', '')),
                    'Applicant Address': str(applicant_row.get('Customer Address', '')),
                    'Similarity Score': score
                })
    
    # Display Registration Number results
    if reg_mismatches:
        st.error(f"Found {len(reg_mismatches)} Registration Number mismatches")
        reg_df = pd.DataFrame(reg_mismatches)
        st.dataframe(reg_df, use_container_width=True)
    else:
        st.success("✅ No Registration Number mismatches found")
    
    st.markdown("---")
    
    st.subheader("📅 RC Date Validation")
    # Display RC Date results
    if date_mismatches:
        st.error(f"Found {len(date_mismatches)} RC Date mismatches")
        date_df = pd.DataFrame(date_mismatches)
        st.dataframe(date_df, use_container_width=True)
    else:
        st.success("✅ No RC Date mismatches found")
    
    st.markdown("---")
    
    st.subheader("🔧 Engine Number Validation")
    # Display Engine Number results
    if engine_mismatches:
        st.error(f"Found {len(engine_mismatches)} Engine Number mismatches")
        engine_df = pd.DataFrame(engine_mismatches)
        st.dataframe(engine_df, use_container_width=True)
    else:
        st.success("✅ No Engine Number mismatches found")
    
    st.markdown("---")
    
    st.subheader("🚗 Chassis Number Validation")
    # Display Chassis Number results
    if chassis_mismatches:
        st.error(f"Found {len(chassis_mismatches)} Chassis Number mismatches")
        chassis_df = pd.DataFrame(chassis_mismatches)
        st.dataframe(chassis_df, use_container_width=True)
    else:
        st.success("✅ No Chassis Number mismatches found")
    
    st.markdown("---")
    
    st.subheader("🏠 Address Validation")
    # Display Address results
    if address_mismatches:
        st.error(f"Found {len(address_mismatches)} Address mismatches")
        address_df = pd.DataFrame(address_mismatches)
        st.dataframe(address_df, use_container_width=True)
    else:
        st.success("✅ No Address mismatches found")
    
    return reg_mismatches + date_mismatches + engine_mismatches + chassis_mismatches + address_mismatches

def create_final_dataframe(data_dict):
    """Create final consolidated dataframe"""
    asset_df = data_dict['asset_df']
    applicant_details_df = data_dict['applicant_details_df']
    aadhar_df = data_dict['aadhar_df']
    pancard_df = data_dict['pancard_df']
    application_form_df = data_dict['application_form_df']
    
    # Get borrowers only
    borrowers_df = applicant_details_df[
        applicant_details_df['Applicant category(borrower/co-borrower/guarantor)'].str.lower() == 'borrower'
    ]
    
    # Merge borrower Aadhaar info
    borrower_aadhar = pd.merge(
        borrowers_df[['Loan Number', 'Aadhar Number']],
        aadhar_df[['Loan Number', 'Aadhar Number', 'age_range', 'state']],
        on=['Loan Number', 'Aadhar Number'],
        how='left'
    )
    
    # Merge PAN info
    borrower_pan = pd.merge(
        borrowers_df[['Loan Number', 'Pancard Number', 'DOB']],
        pancard_df[['Loan Number', 'Pancard Number', 'full_name', 'DOB']],
        on=['Loan Number', 'Pancard Number'],
        how='left'
    )
    
    # Merge Aadhaar and PAN info
    borrower_identity_df = pd.merge(borrower_aadhar, borrower_pan, on='Loan Number', how='outer')
    
    # Trim asset_df
    asset_columns = ['Loan Number', 'rc_number', 'registration_date', 'owner_name', 'mobile_number',
                     'vehicle_category', 'vehicle_chasi_number', 'maker_description', 'maker_model',
                     'color', 'fuel_type', 'manufacturing_date_formatted', 'insurance_company',
                     'insurance_upto', 'permit_number', 'blacklist_status', 'rc_status', 'rto_code']
    
    available_columns = [col for col in asset_columns if col in asset_df.columns]
    asset_df_trimmed = asset_df[available_columns]
    
    # Merge with borrower info
    final_df = pd.merge(asset_df_trimmed, borrower_identity_df, on='Loan Number', how='left')
    
    # Add missing loan numbers from application form
    missing_loan_numbers = application_form_df[~application_form_df['Loan Number'].isin(final_df['Loan Number'])]
    if not missing_loan_numbers.empty:
        missing_rows = pd.DataFrame(columns=final_df.columns)
        missing_rows['Loan Number'] = missing_loan_numbers['Loan Number'].values
        final_df = pd.concat([final_df, missing_rows], ignore_index=True)
    
    return final_df

def create_download_link(df, filename):
    """Create download link for dataframe"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Final_Data', index=False)
    
    st.download_button(
        label="📥 Download Final DataFrame as Excel",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Main application
if api_file and manual_file:
    with st.spinner("Loading and processing data..."):
        data_dict = load_and_process_data(api_file, manual_file)
    
    if data_dict:
        st.success("✅ Data loaded successfully!")
        
        # Create tabs for different validation sections
        tab1, tab2, tab3, tab4 = st.tabs(["👤 Applicant Form", "📋 Application Form", "🚗 Asset Form", "📊 Final Results"])
        
        with tab1:
            st.header("Applicant Form Section Validation")
            st.markdown("### Validating data consistency across Applicant form sections")
            
            # Aadhar Section
            aadhar_mismatches = validate_aadhar_section(
                data_dict['aadhar_df'], 
                data_dict['applicant_details_df']
            )
            
            st.markdown("---")
            
            # Pancard Section
            pancard_mismatches = validate_pancard_section(
                data_dict['pancard_df'], 
                data_dict['applicant_details_df']
            )
            
            st.markdown("---")
            
            # Name Section
            name_mismatches = validate_name_section(
                data_dict['asset_df'], 
                data_dict['applicant_details_df']
            )
            
            st.markdown("---")
            
            # DOB Section
            dob_mismatches = validate_dob_section(
                data_dict['pancard_df'], 
                data_dict['applicant_details_df']
            )
            
            st.markdown("---")
            
            # Mobile Number Section
            mobile_mismatches = validate_mobile_section(
                data_dict['asset_df'], 
                data_dict['applicant_details_df']
            )
        
        with tab2:
            st.header("Application Form Section Validation")
            st.markdown("### Validating data consistency across Application form sections")
            
            # Application Form Validation (Name and Mobile)
            application_mismatches = validate_application_form_section(
                data_dict['asset_df'], 
                data_dict['application_form_df']
            )
        
        with tab3:
            st.header("Asset Form Section Validation")
            st.markdown("### Validating data consistency across Asset form sections")
            
            # Asset Form Validation (Registration, RC Date, Engine, Chassis, Address)
            asset_mismatches = validate_asset_form_section(
                data_dict['asset_df'], 
                data_dict['asset_details_df'],
                data_dict['applicant_details_df'],
                data_dict['application_form_df']
            )
        
        with tab4:
            st.header("Final Results & Download")
            
            # Summary of all mismatches
            total_mismatches = (
                len(aadhar_mismatches) + len(pancard_mismatches) + 
                len(name_mismatches) + len(dob_mismatches) + 
                len(mobile_mismatches) + len(application_mismatches) + 
                len(asset_mismatches)
            )
            
            # Create summary by section
            st.subheader("📋 Validation Summary by Section")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("👤 Applicant Form Issues", 
                         len(aadhar_mismatches) + len(pancard_mismatches) + len(name_mismatches) + len(dob_mismatches) + len(mobile_mismatches))
                st.write("- Aadhar:", len(aadhar_mismatches))
                st.write("- PAN Card:", len(pancard_mismatches))
                st.write("- Name:", len(name_mismatches))
                st.write("- DOB:", len(dob_mismatches))
                st.write("- Mobile:", len(mobile_mismatches))
            
            with col2:
                st.metric("📋 Application Form Issues", len(application_mismatches))
                st.write("- Name & Mobile validation")
            
            with col3:
                st.metric("🚗 Asset Form Issues", len(asset_mismatches))
                st.write("- Registration, Engine, Chassis, Address validation")
            
            st.markdown("---")
            st.metric("🎯 Total Mismatches Found", total_mismatches)
            
            if total_mismatches == 0:
                st.success("🎉 All validations passed! No mismatches found.")
            else:
                st.warning(f"⚠️ Found {total_mismatches} total mismatches across all sections.")
            
            # Create and display final dataframe
            st.subheader("Final Consolidated DataFrame")
            final_df = create_final_dataframe(data_dict)
            
            st.info(f"Final DataFrame contains {len(final_df)} rows and {len(final_df.columns)} columns")
            st.dataframe(final_df.head(10), use_container_width=True)
            
            # Download button
            create_download_link(final_df, f"Final_Dataframe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            
            # Show column info
            with st.expander("📋 Column Information"):
                col_info = pd.DataFrame({
                    'Column Name': final_df.columns,
                    'Data Type': final_df.dtypes.astype(str),
                    'Non-Null Count': final_df.count(),
                    'Null Count': final_df.isnull().sum()
                })
                st.dataframe(col_info, use_container_width=True)

else:
    st.info("👆 Please upload both Excel files to begin validation")
    
    # Show sample structure
    with st.expander("📖 Expected File Structure"):
        st.markdown("""
        **API Excel File (AssetDetails) should contain:**
        - **Aadhar sheet**: Contains Aadhar verification data
        - **pancard sheet**: Contains PAN card verification data  
        - **Asset sheet**: Contains asset/vehicle details
        
        **Manual Excel File should contain:**
        - **ApplicationForm sheet**: Contains application form data
        - **Applicant sheet**: Contains applicant personal details
        - **Asset sheet**: Contains asset registration details
        
        ### Validation Sections:
        
        **👤 Applicant Form Sections:**
        - Aadhar number validation
        - PAN card number validation
        - Name validation
        - Date of birth validation
        - Mobile number validation
        
        **📋 Application Form Sections:**
        - Name validation
        - Mobile number validation
        
        **🚗 Asset Form Sections:**
        - Registration number validation
        - RC date validation
        - Engine number validation
        - Chassis number validation
        - Address validation
        """)        

# Footer
st.markdown("---")
st.markdown("🔍 **Excel Sheet Validation Tool** - Developed for data consistency validation")