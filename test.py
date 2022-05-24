from utils import *
from widgets import *
import time

mentor_pd = './Data/Goodjob Mentorship Program _ Mentor Form (Responses).xlsx'
mentee_pd = './Data/Goodjob Mentorship Program _ Mentee Form (Responses).xlsx'
mentee_pd_temp = './temp/Goodjob Mentorship Program _ Mentee Form (Responses).xlsx'
lang_file = 'Utils/Languages.txt'


t1 = time.time()
# Process number 1
# assign column = 14
check_df = check_if_assign(df_dir=mentee_pd, col_check=14)
# column for id check
new_df = check_if_new_rec(new_ver_dir=mentee_pd, temp_dir=mentee_pd_temp, col_check=1)

t2 = time.time()
# Process number 2
# [language, new role, theme, location]
params_for_mentee = [5, 9, 10, 12]
mentee_df_new_reduce = reduce_df(df=check_df, params=params_for_mentee, id_col=1)

t3 = time.time()
# Process number 3
# [language, current role, theme, location]
params_for_mentor = [4, 8, 10, 13]
mentor_df = import_xlsx(mentor_pd)
mentor_df_reduce = reduce_df(df=mentor_df, params=params_for_mentor, id_col=1)

t4 = time.time()
# Process number 4
# trans mentee with these new columns and get one mentee
cols_to_trans = [1, 2, 3, 4]
mentee_df_new_reduce_trans = trans_df(df=mentee_df_new_reduce, lang='en', cols=cols_to_trans, trans_col_names=False)
val = 'alisa.shynharova@gmail.com'
mentee_df_new_reduce_trans_one = get_mentee(df=mentee_df_new_reduce_trans, val=val, col=0)

t5 = time.time()
# Process number 5
# Filtering the languages for Mentors and Mentees
lang_mentor_ls_uq, lang_mentor_ls, mentor_df_reduce_fil = filter_by(df=mentor_df_reduce, col=1, f_dir=lang_file)
lang_mentee_ls_uq, lang_mentee_ls, mentee_df_new_reduce_trans_one_fill = filter_by(df=mentee_df_new_reduce_trans_one,
                                                                                   col=1, f_dir=lang_file)


t6 = time.time()
# Process number 6
# Matching Mentors to Mentees based on the language
lang_strong_match_ls = strong_match(df_to_fil=mentee_df_new_reduce_trans_one_fill,
                                    df_from_to=mentor_df_reduce_fil, f_dir=lang_file,
                                    words_uq=lang_mentor_ls_uq, words_ls=lang_mentor_ls,
                                    by_col_to=1, col_to=0, col_from=0)

t7 = time.time()
# Process number 7
# Matching Mentors to Mentees based on the text
text_light_match_ls = light_match(df_to_fil=mentee_df_new_reduce_trans_one,
                                  df_from_to_ls=lang_strong_match_ls, by_cols_to=[3], by_cols_from=[3])

t8 = time.time()
# Process number 8
# Update the dataframe
mentee_df = import_xlsx(mentee_pd)
val_to_fill = 'olga.maksimuk96@gmail.com'
mentee_df_update = update_df(df=mentee_df, val_idx=val, val_to_fill=val_to_fill, user='Rafal Marciniak')

t9 = time.time()
# Process number 9
# Save the dataframe into xlsx
#save_xlsx(df=mentee_df, in_dir=mentee_pd, out_dir=mentee_pd)

t10 = time.time()
time_ls = [t1, t2, t3, t4, t5, t6, t7, t8, t9, t10]
for i in range(1, len(time_ls)):
    print('Process number ' + str(i) + ' took :' + str(time_ls[i] - time_ls[i - 1]))


