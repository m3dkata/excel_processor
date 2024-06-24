import streamlit as st
import pandas as pd
import os

def process_excel_file(file, sheet_name='מפה לביצוע'):
    # Load the Excel file
    sheet1_data = pd.read_excel(file, sheet_name=sheet_name)
    st.write("Original Data:")
    st.write(sheet1_data.head())

    # Identify all action columns dynamically
    action_columns = [col for col in sheet1_data.columns if col.startswith('פעולה')]

    # Filter the relevant columns including station code
    columns_of_interest = ['מקט תחנה'] + action_columns
    relevant_data = sheet1_data[columns_of_interest]
    st.write("Filtered Data (Relevant Columns):")
    st.write(relevant_data.head())

    # Remove rows where all 'פעולה' columns are NaN
    relevant_data = relevant_data.dropna(how='all', subset=action_columns)
    st.write("Filtered Data (After Removing NaNs):")
    st.write(relevant_data.head())

    # Create a list of unique actions to produce
    unique_actions = relevant_data.melt(id_vars=['מקט תחנה'], value_vars=action_columns)
    unique_actions = unique_actions.dropna().drop_duplicates().reset_index(drop=True)
    st.write("Unique Actions:")
    st.write(unique_actions.head())

    unique_actions.columns = ['Station Code', 'Action Type', 'Action']

    # Handle 505 actions: Change 'הסרה' actions in 505 to 'סטריפ ריק'
    unique_actions['Action'] = unique_actions['Action'].apply(lambda x: 'סטריפ ריק' if 'הסרה' in x else x)

    # Remove the string 'הוספה -' from the actions except for 'הוספה ברייל -'
    unique_actions['Action'] = unique_actions['Action'].apply(lambda x: x if 'הוספה ברייל -' in x else x.replace('הוספה -', ''))

    # Split Braille actions into multiple rows
    def split_braille(action):
        if 'הוספה ברייל -' in action:
            return ['ברייל ' + part.strip() for part in action.replace('הוספה ברייל -', '').split(',')]
        else:
            return [action]

    split_actions = unique_actions['Action'].apply(split_braille).explode().reset_index(drop=True)
    st.write("Split Actions:")
    st.write(split_actions.head())

    # Clean Braille actions
    def clean_braille_actions(action):
        if 'הסרה ברייל' in action or 'הסרת ברייל' in action:
            return None  # Remove the action if it involves removing Braille
        if 'הוספה ברייל -' in action:
            return action.replace('הוספה ברייל -', '').strip()  # Remove the prefix for adding Braille
        if 'ברייל' in action:
            return action.replace('ברייל ', '').strip()  # Remove the prefix 'ברייל'
        return action

    split_actions = split_actions.apply(clean_braille_actions).dropna().reset_index(drop=True)
    st.write("Cleaned Actions:")
    st.write(split_actions.head())

    # Categorize and sort actions
    def categorize_action(action):
        if action.isdigit():  # Braille actions (only numbers)
            return 0
        elif any(char.isdigit() for char in action):  # Route actions (numbers mixed with text)
            return 1
        else:  # Other actions
            return 2

    bill_of_quantities = split_actions.value_counts().reset_index()
    bill_of_quantities.columns = ['Action', 'Quantity']
    bill_of_quantities['Category'] = bill_of_quantities['Action'].apply(categorize_action)
    bill_of_quantities = bill_of_quantities.sort_values(by=['Category', 'Action']).drop(columns=['Category']).reset_index(drop=True)
    st.write("Bill of Quantities:")
    st.write(bill_of_quantities.head())

    # Sheet 2: Static, Poly, and Fixture Actions
    # Filter rows where any action column contains 'סטטי', 'פולי', או 'מתקן'
    filtered_data = relevant_data[relevant_data[action_columns].apply(lambda x: x.str.contains('סטטי|פולי|מתקן', case=False, na=False).any(), axis=1)]
    st.write("Filtered Data for Static, Poly, Fixture Actions:")
    st.write(filtered_data.head())

    # Extract relevant columns
    filtered_data_output = filtered_data[['מקט תחנה'] + action_columns]

    # Melt the DataFrame to flatten it and keep only rows with relevant actions
    flattened_data = filtered_data_output.melt(id_vars=['מקט תחנה'], value_vars=action_columns, var_name='עמודת פעולה', value_name='פעולה').dropna().sort_values(by=['מקט תחנה', 'עמודת פעולה'])
    relevant_actions = flattened_data[flattened_data['פעולה'].str.contains('סטטי|פולי|מתקן', case=False)]
    result = relevant_actions[['מקט תחנה', 'פעולה']]
    st.write("Relevant Actions:")
    st.write(result.head())

    # Initialize counters for each type of action
    static_count = result['פעולה'].str.contains('סטטי', case=False).sum()
    poly_count = result['פעולה'].str.contains('פולי', case=False).sum()
    fixture_count = result['פעולה'].str.contains('מתקן', case=False).sum()

    # Create a summary DataFrame
    summary = pd.DataFrame({
        'סוג פעולה': ['סטטי', 'פולי', 'מתקן'],
        'כמות': [static_count, poly_count, fixture_count]
    })
    st.write("Summary:")
    st.write(summary)

    # Sheet 3: Flag Actions
    # Identify rows where any action column contains 'דגל'
    flag_actions = relevant_data[relevant_data[action_columns].apply(lambda row: row.str.contains('דגל', case=False, na=False).any(), axis=1)]
    st.write("Flag Actions:")
    st.write(flag_actions.head())

    # Melt the dataframe to have one action per row
    flag_actions_melted = flag_actions.melt(id_vars=['מקט תחנה'], value_vars=action_columns)
    flag_actions_melted = flag_actions_melted.dropna().reset_index(drop=True)

    # Filter rows to keep only those actions that contain 'דגל'
    flag_actions_filtered = flag_actions_melted[flag_actions_melted['value'].str.contains('דגל', case=False)]
    flag_actions_summary = flag_actions_filtered[['מקט תחנה', 'value']]
    flag_actions_summary.columns = ['Station Code', 'Action']
    st.write("Flag Actions Summary:")
    st.write(flag_actions_summary.head())

    # Sheet 4: Station Head Actions
    # Identify rows where any action column contains 'ראש תחנה'
    station_head_actions = relevant_data[relevant_data[action_columns].apply(lambda row: row.str.contains('ראש תחנה', case=False, na=False).any(), axis=1)]
    st.write("Station Head Actions:")
    st.write(station_head_actions.head())

    # Melt the dataframe to have one action per row
    station_head_actions_melted = station_head_actions.melt(id_vars=['מקט תחנה'], value_vars=action_columns)
    station_head_actions_melted = station_head_actions_melted.dropna().reset_index(drop=True)

    # Filter rows to keep only those actions that contain 'ראש תחנה'
    station_head_actions_filtered = station_head_actions_melted[station_head_actions_melted['value'].str.contains('ראש תחנה', case=False)]
    station_head_actions_summary = station_head_actions_filtered[['מקט תחנה', 'value']]
    station_head_actions_summary.columns = ['Station Code', 'Action']
    st.write("Station Head Actions Summary:")
    st.write(station_head_actions_summary.head())

    # Save all sheets to a new Excel file
    output_file_path = 'output/{}'.format(file.name.replace('.xlsx', ' כתב כמויות מלא.xlsx'))
    with pd.ExcelWriter(output_file_path) as writer:
        bill_of_quantities.to_excel(writer, sheet_name='Final Sorted Bill of Quantities', index=False)
        result.to_excel(writer, sheet_name='Static Poly Fixture Actions', index=False)
        summary.to_excel(writer, sheet_name='Static Poly Fixture Summary', index=False)
        flag_actions_summary.to_excel(writer, sheet_name='Flag Actions', index=False)
        station_head_actions_summary.to_excel(writer, sheet_name='Station Head Actions', index=False)

    return output_file_path

def main():
    st.title("Excel File Processor")
    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)
    sheet_name = st.text_input("Enter the sheet name", 'מפה לביצוע')

    if st.button("Process Files"):
        if uploaded_files:
            for uploaded_file in uploaded_files:
                output_file_path = process_excel_file(uploaded_file, sheet_name)
                st.success(f"Processed file saved as {output_file_path}")
                with open(output_file_path, "rb") as file:
                    btn = st.download_button(
                        label="Download Processed File",
                        data=file,
                        file_name=output_file_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning("Please upload at least one Excel file.")

if __name__ == "__main__":
    if not os.path.exists('output'):
        os.makedirs('output')
    main()
