import pandas as pd
import numpy as np


dict_test = {
    'Toan': 'Dang',
    'Anh': 'Van',
    'Van': 'Nam',
    'Khoa hoc': 'Thu',
}

dict_count = {
    'Toan': 8,
    'Anh': 7,
    'Van': 6,
    'Khoa hoc': 2
}

def check_consecutive_lessons(group):
    sorted_buoi_hoc = group.sort_values()
    return (sorted_buoi_hoc.diff() == 1).any()

def check_condition(df, df_2):
    if not df_2.empty:
        df = pd.concat([df, df_2], ignore_index=True)
    # Optimized check for no teacher teaching at the same time
    if df.dropna(subset=['Giáo viên']).duplicated(subset=['Giáo viên', 'Ngày học', 'Giờ học']).any():
        return False

    # Check for no consecutive lessons in the same class and subject
    grouped = df.dropna(subset=['Môn']).groupby(['Lớp', 'Môn'])
    if grouped['Buổi học'].apply(check_consecutive_lessons).any():
        return False

    return True


def assign_teacher(df, subject, result, time):
    count_break = dict_count[subject]
    original_df = df.copy()
    
    total_tries = 2*(time+1)
    n = 0
    filled = False
    
    while n < total_tries and not filled:
        successful_assignments = 0
        indices_to_try = set(df[df['Môn'].isna()].index.tolist())
        while successful_assignments < count_break and indices_to_try:
            total_values = len(original_df[original_df['Môn'] == subject])
            if total_values >= count_break:
                break
            index = np.random.choice(list(indices_to_try), replace=False)
            original_df.loc[index, 'Giáo viên'] = dict_test[subject]
            original_df.loc[index, 'Môn'] = subject
            if check_condition(original_df, df_2=result):
                successful_assignments += 1
            else:
                original_df.loc[index, ['Giáo viên', 'Môn']] = np.nan
            indices_to_try.remove(index)
        filled = successful_assignments >= count_break
        n += 1
        
    return original_df, filled



file_name = input('Enter Excel_file name: ')
df = pd.read_excel(file_name)

indices_to_drop = df[df['Buổi học'] == 24].index
# Drop these indices
df = df.drop(indices_to_drop)
df = df.reset_index(drop=True)

df['Giáo viên'] = df['Giáo viên'].astype('object')
df['Môn'] = df['Môn'].astype('object')



total_loops = 5
for _ in range(total_loops):
    result = pd.DataFrame()
    all_class = np.unique(df['Lớp'].values)
    np.random.shuffle(all_class)
    for time, class_room in enumerate(all_class):
        print(f"Processing classroom: {class_room}")
        df_temp = df[df['Lớp'] == class_room]
        index = 0
        count = 1
        max_attempts = 2*(time+1)
        subjects = list(dict_test.keys())
        np.random.shuffle(subjects)
        while index < len(subjects) and count <= max_attempts:
            df_temp, filled = assign_teacher(df_temp, subjects[index], result=result, time=time)
            if df_temp['Môn'].isna().sum() == 0:
                break
            if filled:
                index = 0 if index == len(subjects) - 1 else index + 1
            else:
                index = max(0, index - 1)
            count += 1
        result = pd.concat([result, df_temp], ignore_index=True)
        time += 1
        
    if result['Môn'].isna().sum() <= 0:
        print("Assignment complete")
        break
    else:
        print("Some classes were not fully assigned, retrying...") 

    
base_name = file_name.split('.')[0]
output_file_name = f"{base_name}_output.xlsx"
result.to_excel(output_file_name)
