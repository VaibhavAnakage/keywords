import pandas as pd
import re
from keyword_definitions import groups  # Import keyword groups from a separate file

# Load Excel file to process
input_file = r".\InputFile.xlsx"
output_file = 'processOutput.xlsx'

try:
    df = pd.read_excel(input_file, sheet_name='Data')
except FileNotFoundError:
    print(f"Error: File '{input_file}' not found.")
    exit()

# Add space at the beginning and end of each entry in 'Description' and 'Resolution'
df['Description'] = df['Description'].apply(lambda x: f' {x} ' if pd.notna(x) else x)
df['Resolution'] = df['Resolution'].apply(lambda x: f' {x} ' if pd.notna(x) else x)

# Function to check for keyword matches
def check_keywords(text, group):
    main_keywords = group['MainKeywords']
    group_keywords = [group[key] for key in group if key.startswith('GroupKeywords')]
    exclusion_patterns = group.get('ExclusionPatterns', [])

    # Convert text to lowercase
    text_lower = text.lower()

    # Skip if any exclusion pattern matches
    for pattern in exclusion_patterns:
        if re.search(pattern, text_lower):
            return None

    matched_keywords = []

    # Function to create a regex pattern for keywords considering spaces
    def keyword_pattern(keyword):
        return re.escape(keyword.lower())

    # Check for main keywords
    for keyword in main_keywords:
        if re.search(keyword_pattern(keyword), text_lower):
            matched_keywords.append(keyword)
        else:
            return None  # If any main keyword doesn't match, skip

    # Check at least one keyword from each GroupKeywords set
    for keyword_set in group_keywords:
        group_matched = False
        for keyword in keyword_set:
            if re.search(keyword_pattern(keyword), text_lower):
                matched_keywords.append(keyword)
                group_matched = True
                break
        if not group_matched:
            return None

    return {
        'GroupDescription': group['GroupDescription'],
        'SubGroupDescription': group['SubGroupDescription'],
        'Addressable': group['Addressable'],
        'Priority-Type': group.get('Priority-Type', [''])[0],
        'Matched Keywords': ', '.join(matched_keywords)
    }

# Separate groups by priority
pro_priority_groups = [g for g in groups if g.get('Priority-Type', [''])[0] == 'Pro']
high_priority_groups = [g for g in groups if g.get('Priority-Type', [''])[0] == 'High']
low_priority_groups = [g for g in groups if g not in pro_priority_groups + high_priority_groups]

# Initialize columns if they don't exist
for col in ['GroupDescription', 'SubGroupDescription', 'Addressable', 'Priority-Type', 'Matched Keywords']:
    if col not in df.columns:
        df[col] = ''

# Iterate through rows and apply matching logic
for index, row in df.iterrows():
    if df.at[index, 'GroupDescription'] == '':  # Only process empty entries
        description = str(row['Description'])
        resolution = str(row['Resolution'])

        matched = False

        # Check Pro Priority
        for group in pro_priority_groups:
            result = check_keywords(description, group) or check_keywords(resolution, group)
            if result:
                for key in result:
                    df.at[index, key] = result[key]
                matched = True
                break

        # Check High Priority
        if not matched:
            for group in high_priority_groups:
                result = check_keywords(description, group) or check_keywords(resolution, group)
                if result:
                    for key in result:
                        df.at[index, key] = result[key]
                    matched = True
                    break

        # Check Low Priority
        if not matched:
            for group in low_priority_groups:
                result = check_keywords(description, group) or check_keywords(resolution, group)
                if result:
                    for key in result:
                        df.at[index, key] = result[key]
                    break

# Save updated data back to Excel
df.to_excel(output_file, index=False)
print(f"Output saved to '{output_file}'.")
