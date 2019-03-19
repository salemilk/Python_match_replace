import numpy as np
import pandas as pd
import warnings

warnings.filterwarnings("ignore")

#读取文件
match = pd.read_excel('match.xlsx', sheet_name='Sheet1')
replace = pd.read_excel('replace.xlsx', sheet_name='Sheet1') #替换文本
sample_input = pd.read_excel('sample_input.xlsx', sheet_name='Sheet1') #源文件

#读取需要处理的列
match_col_n1 = ['user_id', 'task_name', 'subtask_name', 'comments', 'marketplace', 'week', 'label'] #条件
match_col_n2 = ['user_id', 'task_name', 'subtask_name', 'comments', 'marketplace', 'week'] #条件
match_col_n3 = ['task_name', 'subtask_name', 'marketplace']

replace_col_n1 = ['task_name', 'subtask_name', 'comments', 'marketplace', 'label'] #替换文本
replace_col_n2 = ['task_name', 'subtask_name', 'marketplace', 'label'] #替换文本
sample_input_col_n = ['DATE', 'start_date_local', 'end_date_local', 'log_time', 'user_id', 'node_id',
                      'productive', 'program_short_name', 'task_name', 'subtask_name', 'comments', 'marketplace',
                      'manager_id', 'week'] #元文件


#DataFrame
match_data = pd.DataFrame(match,  columns=match_col_n1)
match_data[['label', 'week']] = match_data[['label', 'week']].astype(np.float)

match_data1 = match_data[~pd.isnull(match_data['user_id'])]
match_data2 = match_data[pd.isnull(match_data['user_id'])]

match_data1 = match_data1.astype(str)
match_data2 = match_data2.astype(str)

replace_data = pd.DataFrame(replace, columns=replace_col_n1)
replace_data['label'] = replace_data['label'].astype(np.float)
replace_data = replace_data.astype(str)

sample_input_data = pd.DataFrame(sample_input, columns=sample_input_col_n).astype(str)
sample_input_data['week'] = sample_input_data['week'].astype(np.float)
sample_input_data = sample_input_data.astype(str)


# 合并两张表之后的临时表
tmp1 = sample_input_data.merge(match_data1, how='left', on=None, 
                 left_on=match_col_n2, right_on=match_col_n2, 
                 left_index=False, right_index=False, sort=False, 
                 suffixes=('_x', '_y'), copy=True, 
                 indicator=False, validate=None)

# 合并两张表之后的临时表
tmp2 = sample_input_data.merge(match_data2, how='left', on=None, 
                 left_on=match_col_n3, right_on=match_col_n3, 
                 left_index=False, right_index=False, sort=False, 
                 suffixes=('_x', '_y'), copy=True, 
                 indicator=False, validate=None)


index1 = tmp1[~pd.isnull(tmp1['label'])].index.values
index2 = tmp2[~pd.isnull(tmp2['label'])].index.values
index2 = np.array(list(set(index2) - set(index1)))

sample_input_data.loc[index1, 'label'] = tmp1['label']
sample_input_data.loc[index2, 'label'] = tmp2['label']

sample_input_data['comments'] = sample_input_data['comments'].replace('nan', '')

for num in replace_data.index.values:
    label = replace_data.loc[num, 'label']
    replace4_index = sample_input_data.loc[index1, :][sample_input_data['label'] == label].index
    replace3_index = sample_input_data.loc[index2, :][sample_input_data['label'] == label].index
    sample_input_data.loc[replace4_index, replace_col_n1] = replace_data.loc[num:num, replace_col_n1].values
    sample_input_data.loc[replace3_index, replace_col_n2] = replace_data.loc[num:num, replace_col_n2].values

sample_input_data.to_excel('output.xlsx', index=False)