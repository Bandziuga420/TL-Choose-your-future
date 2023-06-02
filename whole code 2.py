import pandas as pd


QA_file = r'C:\Users\Aleks\Documents\University\LUISS - Nova SBE Master\Nova SBE\TechLabs (Data Science)\Project\Questions_answers.xlsx'
Stat_file = r'C:\Users\Aleks\Documents\University\LUISS - Nova SBE Master\Nova SBE\TechLabs (Data Science)\Project\Sample_statements.xlsx'
Output_file = r'C:\Users\Aleks\Documents\University\LUISS - Nova SBE Master\Nova SBE\TechLabs (Data Science)\Project\Output.xlsx'



# Read the Questions_answers
QA = pd.read_excel(QA_file)

# Get the statements from column Statements
statements = QA['Statements'].tolist()

# Ask the user for input and register it in the "Answers" column
answers = []
for statement in statements:
    print(statement)
    user_input = input("Enter 1 for Agree, 0 for Disagree: ")
    answers.append(user_input)

# Update the "Answers" column in the DataFrame
QA['Answers'] = answers

# Save the updated DataFrame to the same Questions_answers file
QA.to_excel(QA_file, index=False)

# Read the updated file
QA_updated = pd.read_excel(QA_file)

print("------------------------- File updated ----------------------")




# Read the Questions_with_answers & Sample_statements
QA = pd.read_excel(QA_file)
Stat = pd.read_excel(Stat_file)

# Merge the two dataframes based on the 'Statements' column
merged_df = pd.merge(QA, Stat, on='Statements', how='left')

# Create a new dataframe to store the filtered results
output_data = []

# Iterate over each row and check the 'Answers' column
for index, row in merged_df.iterrows():
    statement = row['Statements']
    answer = row['Answers']

# Check if the answer is 1
    if answer == 1:                          
        output_row = [
            statement,
            row['Business and Economics'],
            row['Technology and Science'],
            row['Literature and Humanities'],
            row['Arts and Design'],
            row['Social Sciences'],
            row['Languages and Linguistics'],
            row['History and Geography']
        ]
        output_data.append(output_row)

# Create a DataFrame from the output data
output_df = pd.DataFrame(output_data, columns=['Statements',
                                               'Business and Economics',
                                               'Technology and Science',
                                               'Literature and Humanities',
                                               'Arts and Design',
                                               'Social Sciences',
                                               'Languages and Linguistics',
                                               'History and Geography'])

# Save the results to the Output
output_df.to_excel(Output_file, index=False)

print("------------------------- Output saved ----------------------")




# Read the Output file
outp = pd.read_excel(Output_file)

# Count the letters in each column
count_b = outp['Business and Economics'].apply(lambda x: str(x).count('b') if isinstance(x, str) else 0).sum()
count_c = outp['Technology and Science'].apply(lambda x: str(x).count('c') if isinstance(x, str) else 0).sum()
count_d = outp['Literature and Humanities'].apply(lambda x: str(x).count('d') if isinstance(x, str) else 0).sum()
count_e = outp['Arts and Design'].apply(lambda x: str(x).count('e') if isinstance(x, str) else 0).sum()
count_f = outp['Social Sciences'].apply(lambda x: str(x).count('f') if isinstance(x, str) else 0).sum()
count_g = outp['Languages and Linguistics'].apply(lambda x: str(x).count('g') if isinstance(x, str) else 0).sum()
count_h = outp['History and Geography'].apply(lambda x: str(x).count('h') if isinstance(x, str) else 0).sum()

# Print the counts for each column
print()
print("RESULTS")
print()

print("Business and Economics score is:", count_b)
print("Technology and Science score is:", count_c)
print("Literature and Humanities score is:", count_d)
print("Arts and Design score is:", count_e)
print("Social Sciences score is:", count_f)
print("Languages and Linguistics score is:", count_g)
print("History and Geography score is:", count_h)
