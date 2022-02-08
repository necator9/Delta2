import os
import subprocess
import pandas as pd
import numpy as np
import difflib
import functools		


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
	hid_format = workbook.add_format({'text_wrap': False})
	for col in cols:
		col_idx = df.columns.get_loc(col)
		worksheet.set_column(col_idx + 1, col_idx + 1, None, hid_format, {'hidden': True})
		row_idxs = [out_df.index.get_loc(i) for i in out_df[pd.notnull(out_df[col])].index.tolist()]
		for row in row_idxs:
			worksheet.write(row + 1, col_idx + 1, out_df.iloc[row, col_idx], hid_format)



def highlight_diff_cells(df_out, changed, worksheet):
	# Highlight entire rows
	formatters = {'New': workbook.add_format({'bg_color': '#cae3d3', 'border': 3, 'text_wrap': True}),
			 	  'Changed': workbook.add_format({'bg_color': '#f7e6cb', 'border': 3, 'text_wrap': True}),
			      'Deleted': workbook.add_format({'bg_color': '#ff8a82', 'border': 3, 'text_wrap': True})}

	for typ, form in formatters.items():
		row_idxs = [df_out.index.get_loc(i) for i in df_out[df_out['Delta'] == typ].index.tolist()]
		for row in row_idxs:
			worksheet.set_row(row + 1, None, form)

	# Highlight changed cells only
	changed_bright = workbook.add_format({'bg_color': '#ffd399', 'border': 3, 'text_wrap': True})

	for col_name, ch_df in changed.items():
		col_idx = df_out.columns.get_loc(col_name)
		row_idxs = [df_out.index.get_loc(i) for i in ch_df[ch_df['Delta'] == 'Changed'].index.tolist()]
		for row in row_idxs:
			worksheet.write(row + 1, col_idx + 1, df_out.iloc[row, col_idx], changed_bright)



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


def preprocess(df, index_col):
	# Remove special characters (CB export handled)
	def remove_sp_ch(df):
		cols_to_filter = ['Description']
		specs_to_remove = ['~', '__', '\'\'', '\\_x000D_', '\\\\', '\\r']
		for col in cols_to_filter:
			df[col] = df[col].replace(specs_to_remove, '', regex=True)
		return df

	# Filter special characters
	df = remove_sp_ch(df)
	# Set index column
	df = df.set_index(index_col)
	# Hadle NaN != Nan
	df = df.replace({np.nan: None}) 

	return df


def find_delta(df_old, df_new, tracked_cols):
	persist = df_new[df_new.index.isin(df_old.index)]
	new = df_new[~df_new.index.isin(df_old.index)]
	deleted = df_old[~df_old.index.isin(df_new.index)]

	changed = dict()
	for col in tracked_cols:
		changed[col] = persist[persist[col].ne(df_old[col])]

	return new, changed, deleted


f_old = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R4_4_PH_CodeBeamer\AVE 22.0-R4.4.xlsx'
f_new = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R5_2_PH_CodeBeamer\AVE 25.0-R5.2.xlsx'

df_old = pd.read_excel(f_old)
df_new = pd.read_excel(f_new)

index_col = 'ID'
df_old, df_new = [preprocess(df, index_col) for df in [df_old, df_new]]

# Columns to track changed requrements
tracked_cols = ['Description', 'Feature']
new, changed, deleted = find_delta(df_old, df_new, tracked_cols)

# Filter out Information and Heading changes
filter_info_heading = True
if filter_info_heading:
	ignoretypes = ['Information', 'Heading']
	changed = {fe_c: changed[fe_c][~changed[fe_c]['Type'].isin(ignoretypes)] for fe_c in tracked_cols}
	new = new[~new['Type'].isin(ignoretypes)]
	deleted = deleted[~deleted['Type'].isin(ignoretypes)]

n_changed = functools.reduce(lambda left, right: pd.merge(left, right, how='outer', left_index=True, right_index=True), 
			list(changed.values())).shape[0]  # Combile all changed columns and get size
n_new = new.shape[0]
n_deleted = deleted.shape[0]
# print(changed[tracked_cols[0]].shape[0], changed[tracked_cols[1]].shape[0])
# print(n_changed, tt.shape[0])

out_df = df_new.copy()

out_df.loc[out_df.index.isin(new.index), 'Delta'] = 'New'
for fe_c in tracked_cols:
	out_df.loc[out_df.index.isin(changed[fe_c].index), 'Delta'] = 'Changed'
	changed[fe_c]['Delta'] = 'Changed'

show_deleted = True
if show_deleted:
	deleted['Delta'] = 'Deleted'
	out_df = pd.concat([out_df, deleted])
else:
	deleted = pd.DataFrame(columns=out_df.columns.tolist()) # Override by empty df

# Use first name given in tracked columns as a main column. Diff is found for this column.
main_diff_col = tracked_cols[0]
for ch_i in changed[main_diff_col].index.tolist():
	out_df.at[ch_i, 'Old Description'] = df_old.loc[ch_i]['Description']


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

# changed = changed['Description'] 

highlight_diff_chars(df_old, df_new, changed[main_diff_col], worksheet)
highlight_diff_cells(out_df, changed, worksheet)
print(changed['Feature'])

beutifulize(worksheet, out_df)		

writer.save()

# Open output file
excel = 'C:\\Program Files (x86)\\Microsoft Office\\Office16\\EXCEL.EXE'
subprocess.call(('cmd', '/c', 'start', '', excel, '', out_file))

# os.system(excel)
# a = '""alox""'
# print(a.replace('\"\"', "1"))"''