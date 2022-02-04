import os
import subprocess
import pandas as pd
import numpy as np
import difflib


# Remove special characters (CB export handled)
def remove_sp_ch(df, cols_to_filter):
	specs_to_remove = ['~', '__', '\'\'', '\\_x000D_', '\\\\', '\\r']
	for col in cols_to_filter:
		df[col] = df[col].replace(specs_to_remove, '', regex=True)

	return df


def beutifulize(worksheet, df):
	worksheet.freeze_panes(1, 0)

	# Apply autofiltering (drop-down excel menu) for selected columns
	(max_row, max_col) = df.shape
	worksheet.autofilter(0, 0, max_row, max_col)

	headings_idx = df.index[df['Type'] == 'Heading']
	headings_idx = [df.index.get_loc(i) + 1 for i in headings_idx]
	heading_format = workbook.add_format({'bold': True})
	for i in headings_idx:
		worksheet.set_row(i, None, heading_format)

	info_idx = df.index[df['Type'] == 'Information'].tolist()
	info_idx = [df.index.get_loc(i) + 1 for i in info_idx]
	info_format = workbook.add_format({'italic': True})
	for i in info_idx:
		worksheet.set_row(i, None, info_format)

	# Format columns width automatically
	cols = ['ID', 'Object ID_FKT_LAH_DOORS', 'Object Identifier', 'Feature', 'Status', 'Umsetzung', 'Type']
	cols_idx = [df.columns.get_loc(c) + 1 for c in cols if c in df]
	lengths = [max([len(str(s)) for s in df[col].values]) for col in cols if col in df]
	for length, idx in zip(lengths, cols_idx):
		worksheet.set_column(idx, idx, length + 1, None, None)

	# Format columns width manually
	cols = ['Descr replSigs', 'Description', 'Diff' ]
	desc_format = workbook.add_format({'text_wrap': True})
	cols_idx = [df.columns.get_loc(c) + 1 for c in cols if c in df]
	length = 60
	for idx in cols_idx:
		worksheet.set_column(idx, idx, length, desc_format, None)

	# # Hide unused columns
	cols = ['Modified at', 'CR Referenz', 'Old Description']
	cols_idx = [df.columns.get_loc(c) + 1 for c in cols if c in df]
	hid_format = workbook.add_format({'text_wrap': False, 'font_color': 'red'})
	for idx in cols_idx:
		worksheet.set_column(idx, idx, None, hid_format, {'hidden': False})


def highlight_diff(df_out, new, changed, deleted, worksheet):
	def highlight_row(worksheet, r_idx, formatter):
		for i in r_idx:
			worksheet.set_row(i, None, formatter)

	ch_format = workbook.add_format({'bg_color': '#ffd399', 'text_wrap': True, 'border': 3})
	changed_idx = [df_out.index.get_loc(i) + 1 for i in changed.index.tolist()]
	highlight_row(worksheet, changed_idx, ch_format)

	new_format = workbook.add_format({'bg_color': '#cae3d3', 'text_wrap': True, 'border': 3})
	new_idx = [df_out.index.get_loc(i) + 1 for i in new.index.tolist()]
	highlight_row(worksheet, new_idx, new_format)

	del_format = workbook.add_format({'bg_color': '#ff8a82', 'text_wrap': True, 'border': 3})
	del_idx = [df_out.index.get_loc(i) + 1 for i in deleted.index.tolist()]
	highlight_row(worksheet, del_idx, del_format)


def find_diff(text1, text2):
	def insert_el(diff, add_s, rem_s):
		i = 0
		higher_limit = len(diff)
		while i < higher_limit:
			val = diff[i]
			if val.startswith('+'):
				diff.insert(i, add_s)
				higher_limit += 1
				i += 1
			if val.startswith('-'):
				diff.insert(i, rem_s)
				higher_limit += 1
				i += 1
			i += 1

	diff = list(difflib.unified_diff(text1, text2, n=5000))[3:]

	add_form = workbook.add_format({'bold': True, 'font_color': 'green', 'num_format': '@', 'font_size': 12})
	rem_form = workbook.add_format({'bold': True, 'font_color': 'red', 'num_format': '@', 'font_size': 12})

	insert_el(diff, add_form, rem_form)

	diff = [el[1] if type(el) == str and len(el) == 2 and el[0] in ['+', '-'] else el for el in diff]
	diff = [el.replace(' ', '', 1) if type(el) == str and len(el) == 2 else el for el in diff]


	return diff


def highlight_diff_chars(df_old, df_new, changed, worksheet):
	for ch_i in changed.index.tolist():
		old_text = df_old.loc[ch_i]['Description']
		new_text = df_new.loc[ch_i]['Description']
		diff = find_diff(old_text, new_text)
		worksheet.write_rich_string(out_df.index.get_loc(ch_i) + 1, out_df.columns.get_loc('Diff') + 1, *diff)


f_old = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R4_4_PH_CodeBeamer\AVE 22.0-R4.4.xlsx'
f_new = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R5_2_PH_CodeBeamer\AVE 25.0-R5.2.xlsx'

cols_to_filter = ['Description']
df_old = remove_sp_ch(pd.read_excel(f_old), cols_to_filter)
df_new = remove_sp_ch(pd.read_excel(f_new), cols_to_filter)

df_old = df_old.astype({'ID': int})
df_new = df_new.astype({'ID': int})
df_old = df_old.set_index('ID')
df_new = df_new.set_index('ID')

# Hadle NaN != Nan
df_old = df_old.replace({np.nan: None})
df_new = df_new.replace({np.nan: None})

persist = df_new[df_new.index.isin(df_old.index)]

changed = persist[persist['Description'].ne(df_old['Description'])]
new = df_new[~df_new.index.isin(df_old.index)]
deleted = df_old[~df_old.index.isin(df_new.index)]

# Filter out Information and Heading changes
filter_info_heading = True
if filter_info_heading:
	ignoretypes = ['Information', 'Heading']
	changed = changed[~changed['Type'].isin(ignoretypes)]
	new = new[~new['Type'].isin(ignoretypes)]
	deleted = deleted[~deleted['Type'].isin(ignoretypes)]

n_changed = changed.shape[0]
n_new = new.shape[0]
n_deleted = deleted.shape[0]

out_df = df_new.copy()

out_df.loc[out_df.index.isin(new.index), 'Delta'] = 'New'
out_df.loc[out_df.index.isin(changed.index), 'Delta'] = 'Changed'

show_deleted = True
if show_deleted:
	deleted['Delta'] = 'Deleted'
	out_df = pd.concat([out_df, deleted])
else:
	deleted = pd.DataFrame(columns=out_df.columns.tolist()) # Override by empty df




for ch_i in changed.index.tolist():
	old_text = df_old.loc[ch_i]['Description']
	out_df.at[ch_i, 'Old Description'] = old_text
	# out_df.at[ch_i, 'Diff'] = ''.join(old_text)


col = out_df.pop('Description')
out_df.insert(4, col.name, col) # Ins 3. col Index einfügen
col = out_df.pop('Old Description')
out_df.insert(5, col.name, col)
out_df['Diff'] = np.nan
col = out_df.pop('Diff')
out_df.insert(6, col.name, col)


out_file = r'C:\opt\diff_out.xlsx'
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
out_df.to_excel(writer, sheet_name='diff', index=True)

workbook = writer.book
worksheet = writer.sheets['diff']


# highlight_diff(out_df, new, changed, deleted, worksheet)
# beutifulize(worksheet, out_df)
highlight_diff_chars(df_old, df_new, changed, worksheet)
writer.save()

# Open output file
excel = 'C:\\Program Files (x86)\\Microsoft Office\\Office16\\EXCEL.EXE'
subprocess.call(('cmd', '/c', 'start', '', excel, '', out_file))

# os.system(excel)
# a = '""alox""'
# print(a.replace('\"\"', "1"))"''