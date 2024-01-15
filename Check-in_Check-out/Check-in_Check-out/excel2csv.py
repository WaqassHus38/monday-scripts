import pandas as pd

checkin_actions = ['RELOCATE', 'TO BE CANCELLED', 'SEND CHECKIN', 'DELAYED CHECKIN']

df = pd.read_excel('organized_data.xlsx')

df['action'] = df['action'].str.strip()
df['group'] = df['group'].str.strip()
df['action'] = df['action'].replace('SEND CHECKOUT INSTRUCTIONS', 'SEND CHECKOUT')

errors_df = df[df['action'] == 'URGENT - CLASH']
errors_str = errors_df.to_csv(index=False)

# checkins_df = df[df['action'].isin(checkin_actions)]
checkins_df = df[(df['action'].isin(checkin_actions)) |
                 ((df['action'] == 'URGENT - Early check-in & Late check-out') &
                  (df['group'] == '1 Week Before Check In'))]
checkins_str = checkins_df.to_csv(index=False)

# checkouts_df = df[df['action'] == 'SEND CHECKOUT']
checkouts_df = df[(df['action'] == 'SEND CHECKOUT') |
                  ((df['action'] == 'NO CLASH') & (df['group'] == 'Check Outs'))]
checkouts_str = checkouts_df.to_csv(index=False)

head1_str = "ERRORS,,,,,,,,,"
head2_str = ",,,,,,,,,\nCheckins,,,,,,,,,"
head3_str = ",,,,,,,,,\n,,,,,,,,,\n,,,,,,,,,\nCheckouts,,,,,,,,,"

combined_str = head1_str + '\n' + errors_str + head2_str \
               + '\n' + checkins_str + head3_str + '\n' \
               + checkouts_str

with open('action.csv', 'w', newline='') as f:
    f.write(combined_str)
