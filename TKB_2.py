import pandas as pd
import numpy as np


file_name = input('Enter Excel_file name: ')
df = pd.read_excel(file_name)

indices_to_drop = df[df['Buổi học'] == 24].index
# Drop these indices
df = df.drop(indices_to_drop)
df = df.reset_index(drop=True)

df['Giáo viên'] = df['Giáo viên'].astype('object')
df['Môn'] = df['Môn'].astype('object')


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

def create_subs_list():
    subjects = list(dict_test.keys())
    np.random.shuffle(subjects)
    result = []
    for sub in subjects:
        for _ in range(dict_count[sub]):
            result.append(sub)
    np.random.shuffle(result)
    return result


def assign_teacher(df, result):

    original_df = df.copy()
    subjects = create_subs_list()
    index_sub = range(len(subjects))
    indices_to_try = set(df[df['Môn'].isna()].index.tolist())
    n = 0
    n_tries = 3
    while n < n_tries:
        done_sub = []
        np.random.shuffle(subjects)
        for index in indices_to_try:
            for i in index_sub:    
                if i in done_sub:
                    continue
                original_df.loc[index, 'Giáo viên'] = dict_test[subjects[i]]
                original_df.loc[index, 'Môn'] = subjects[i]
                if check_condition(original_df, df_2=result):
                    done_sub.append(i)
                    break
                else:
                    original_df.loc[index, ['Giáo viên', 'Môn']] = np.nan
        if original_df['Môn'].isna().sum() == 0:
            break
        else:
            original_df.loc[list(indices_to_try), ['Giáo viên', 'Môn']] = np.nan
        n += 1
    return original_df


n = 0
tries = 500
while n < tries:
    result = pd.DataFrame()
    all_class = np.unique(df['Lớp'].values)
    np.random.shuffle(all_class)
    for class_room in all_class:
        print(f"Processing classroom: {class_room}")
        df_temp = df[df['Lớp'] == class_room]
        df_temp = assign_teacher(df_temp, result)
        result = pd.concat([result, df_temp], ignore_index=True)
    if result['Môn'].isna().sum() == 0:
        break
    else:
        print('Retrying...')
    n += 1

if n == tries:
    print('Some classes were not fully assigned!!')

base_name = file_name.split('.')[0]
output_file_name = f"{base_name}_output_2.xlsx"
result.to_excel(output_file_name)
