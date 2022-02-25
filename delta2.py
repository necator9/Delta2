import os
import subprocess
import pandas as pd
import numpy as np
import difflib
import functools
import warnings

warnings.filterwarnings("ignore")   


class Delta2Exception(Exception):
    pass


class Delta2(object):
    def __init__(self, f_old, f_new, out_file, index_col, tracked_cols, filter_info_heading):
        self.f_old = f_old
        self.f_new = f_new
        self.index_col = index_col
        self.tracked_cols = tracked_cols
        self.pd_diff = None
        self.out_file = out_file
        self.main_diff_col = tracked_cols[0]   # Use first name given in tracked columns as a main column. Diff is found for this column.
        self.filter_info_heading = filter_info_heading

    def process(self):
        self.pd_diff = PandasDiffHandler(self.f_old, self.f_new, self.index_col, self.tracked_cols)
        out_df, changed, stats = self.pd_diff.find_pd_diff(filter_info_heading=self.filter_info_heading)

        excel_diff = ExcelDiffHandler(self.out_file, out_df, self.main_diff_col)
        excel_diff.format_exel_output(self.pd_diff.df_old, self.pd_diff.df_new, changed)
        excel_diff.quit()

        return *stats, sum(stats)  # (n_new, n_changed, n_deleted)


class PandasDiffHandler(object):
    def __init__(self, f_old, f_new, index_col, tracked_cols):
        df_old = pd.read_excel(f_old)
        df_new = pd.read_excel(f_new)

        self.tracked_cols = tracked_cols       # Columns to track changed requrements
        self.main_diff_col = tracked_cols[0]   # Use first name given in tracked columns as a main column. Diff is found for this column.

        self.df_old, self.df_new = [self.preprocess(df, index_col) for df in [df_old, df_new]]
        

    def preprocess(self, df, index_col):
        # Filter special characters
        df = self.remove_sp_ch(df)
        # Set index column
        if index_col in df.columns.tolist():
            df = df.set_index(index_col)
        else:
            raise Delta2Exception(f'Index column "{index_col}"" is not in the dataframe. Check your excel files.' )
        
        # Hadle NaN != Nan
        df = df.replace({np.nan: None}) 

        return df

    # Remove special characters (CB export handled)
    def remove_sp_ch(self, df):
        cols_to_filter = self.tracked_cols
        specs_to_remove = ['~', '__', '\'\'', '\\_x000D_', '\\\\', '\\r', '\$']
        for col in cols_to_filter:
            if col in df.columns.tolist():
                df[col] = df[col].replace(specs_to_remove, '', regex=True)
            else:
                raise Delta2Exception(f'Tracked column "{col}" is not in the dataframe. Check your excel files.' )

        return df


    def find_pd_diff(self, filter_info_heading=True):
        def find_delta(df_old, df_new, tracked_cols):
            persist = df_new[df_new.index.isin(df_old.index)]
            new = df_new[~df_new.index.isin(df_old.index)]
            deleted = df_old[~df_old.index.isin(df_new.index)]

            changed = dict()
            for col in tracked_cols:
                changed[col] = persist[persist[col].ne(df_old[col])]

            return new, changed, deleted

        new, changed, deleted = find_delta(self.df_old, self.df_new, self.tracked_cols)

        # Filter out Information and Heading changes
        if filter_info_heading:
            ignoretypes = ['Information', 'Heading']
            changed = {fe_c: changed[fe_c][~changed[fe_c]['Type'].isin(ignoretypes)] for fe_c in self.tracked_cols}
            new = new[~new['Type'].isin(ignoretypes)]
            deleted = deleted[~deleted['Type'].isin(ignoretypes)]

        n_changed = functools.reduce(lambda left, right: pd.merge(left, right, how='outer', left_index=True, right_index=True), 
                    list(changed.values())).shape[0]  # Combile all changed columns and get size
        n_new = new.shape[0]
        n_deleted = deleted.shape[0]

        out_df = self.df_new.copy()

        out_df.loc[out_df.index.isin(new.index), 'Delta'] = 'New'
        for fe_c in self.tracked_cols:
            out_df.loc[out_df.index.isin(changed[fe_c].index), 'Delta'] = 'Changed'
            changed[fe_c]['Delta'] = 'Changed'

        show_deleted = True
        if show_deleted:
            deleted['Delta'] = 'Deleted'
            out_df = pd.concat([out_df, deleted])
        else:
            deleted = pd.DataFrame(columns=out_df.columns.tolist()) # Override by empty df

        out_df[f'Old {self.main_diff_col}'] = np.nan
        out_df[f'Old {self.main_diff_col}'] = out_df[f'Old {self.main_diff_col}'].astype(str)
        for ch_i in changed[self.main_diff_col].index.tolist():
            out_df.at[ch_i, f'Old {self.main_diff_col}'] = self.df_old.loc[ch_i][self.main_diff_col]

        # Fix order in the output df
        col = out_df.pop(self.main_diff_col)
        out_df.insert(4, col.name, col) # Ins 3. col Index einfügen
        col = out_df.pop(f'Old {self.main_diff_col}')
        out_df.insert(5, col.name, col)
        out_df['Diff'] = np.nan
        col = out_df.pop('Diff')
        out_df.insert(6, col.name, col)

        return out_df, changed, (n_new, n_changed, n_deleted)


class ExcelDiffHandler(object):
    def __init__(self, out_file, out_df, main_diff_col):
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        self.writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
        self.out_df = out_df
        out_df.to_excel(self.writer, sheet_name='diff', index=True)
        self.workbook = self.writer.book
        self.worksheet = self.writer.sheets['diff']
        self.main_diff_col = main_diff_col


    def format_exel_output(self, df_old, df_new, changed):
        self.highlight_diff_chars(df_old, df_new, changed[self.main_diff_col])
        self.highlight_diff_cells(changed)
        self.beutifulize()  


    def highlight_diff_chars(self, df_old, df_new, changed_col):
        out_df = self.out_df
        for ch_i in changed_col.index.tolist():
            old_text = df_old.loc[ch_i][self.main_diff_col]
            new_text = df_new.loc[ch_i][self.main_diff_col]
            diff = self.find_diff(old_text, new_text)
            self.worksheet.write_rich_string(out_df.index.get_loc(ch_i) + 1, out_df.columns.get_loc('Diff') + 1, *diff)


    def find_diff(self, text1, text2):
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

        diff = list(difflib.unified_diff(str(text1), str(text2), n=5000))[3:]

        add_form = self.workbook.add_format({'bold': True, 'font_color': 'green', 'num_format': '@', 'font_size': 12})
        rem_form = self.workbook.add_format({'bold': True, 'font_color': 'red', 'font_strikeout': True, 'num_format': '@', 'font_size': 12})

        insert_el(diff, add_form, rem_form)

        diff = [el[1] if type(el) == str and len(el) == 2 and el[0] in ['+', '-'] else el for el in diff]
        diff = [el.replace(' ', '', 1) if type(el) == str and len(el) == 2 else el for el in diff]

        return diff


    def highlight_diff_cells(self, changed):
        df_out = self.out_df
        # Highlight entire rows
        formatters = {'New': self.workbook.add_format({'bg_color': '#cae3d3', 'border': 3, 'text_wrap': True}),
                      'Changed': self.workbook.add_format({'bg_color': '#f7e6cb', 'border': 3, 'text_wrap': True}),
                      'Deleted': self.workbook.add_format({'bg_color': '#ff8a82', 'border': 3, 'text_wrap': True})}

        for typ, form in formatters.items():
            row_idxs = [df_out.index.get_loc(i) for i in df_out[df_out['Delta'] == typ].index.tolist()]
            for row in row_idxs:
                self.worksheet.set_row(row + 1, None, form)

        # Highlight changed cells only
        changed_bright = self.workbook.add_format({'bg_color': '#ffd399', 'border': 3, 'text_wrap': True})

        for col_name, ch_df in changed.items():
            col_idx = df_out.columns.get_loc(col_name)
            row_idxs = [df_out.index.get_loc(i) for i in ch_df[ch_df['Delta'] == 'Changed'].index.tolist()]
            for row in row_idxs:
                self.worksheet.write(row + 1, col_idx + 1, df_out.iloc[row, col_idx], changed_bright)


    def beutifulize(self):
            df = self.out_df
            workbook = self.workbook
            worksheet = self.worksheet
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
            # cols = df.columns.tolist()
            cols_idx = [df.columns.get_loc(c) + 1 for c in cols if c in df]
            lengths = [max([len(str(s)) for s in df[col].values]) for col in cols if col in df]
            for length, idx in zip(lengths, cols_idx):
                worksheet.set_column(idx, idx, length + 1, None, None)

            # Format columns width manually
            cols = ['Descr replSigs', self.main_diff_col, 'Diff' ]
            desc_format = workbook.add_format({'text_wrap': True})
            cols_idx = [df.columns.get_loc(c) + 1 for c in cols if c in df]
            length = 60
            for idx in cols_idx:
                worksheet.set_column(idx, idx, length, desc_format, None)

            # Hide unused columns
            cols = ['Modified at', 'CR Referenz', f'Old {self.main_diff_col}']
            cols_idx = [df.columns.get_loc(c) + 1 for c in cols if c in df]
            hid_format = workbook.add_format({'text_wrap': False})
            for col in cols:
                if col not in df:
                    continue
                col_idx = df.columns.get_loc(col)
                worksheet.set_column(col_idx + 1, col_idx + 1, None, hid_format, {'hidden': True})
                row_idxs = [df.index.get_loc(i) for i in df[pd.notnull(df[col])].index.tolist()]
                for row in row_idxs:
                    worksheet.write(row + 1, col_idx + 1, df.iloc[row, col_idx], hid_format)


    def quit(self):
        self.writer.save()


class ExcelStatsHandler(object):
    def __init__(self, out_file, out_df):
        self.writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
        self.out_df = out_df
        out_df.to_excel(self.writer, sheet_name='stats', index=False)
        self.workbook = self.writer.book
        self.worksheet = self.writer.sheets['stats']

    def plot_pie(self, name, numbers_col, loc, rows):
        # Chart total
        chart = self.workbook.add_chart({'type': 'pie'})
        module_name_col = 0
        numbers_col = numbers_col
        chart.add_series({'name': name,
                                 'categories': ['stats', 1, module_name_col, rows, module_name_col],
                                 'values':     ['stats', 1, numbers_col, rows, numbers_col],})

        chart.set_title({'name': name})
        chart.set_style(10)
        self.worksheet.insert_chart(loc, chart, {'x_offset': 25, 'y_offset': 10})

    def format_exel_output(self, rows):
        pies = [['Total', 4, 'H2'], ['New', 1, 'H17'], ['Changed', 2, 'P2'], ['Deleted', 3, 'P17']]
        for name, numbers_col, loc in pies:
            self.plot_pie(name, numbers_col, loc, rows)

    def quit(self):
        self.writer.save()


if __name__ == '__main__':
    f_old_path = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R4_4_PH_CodeBeamer\LBS 19.0-R4.4.xlsx'
    # f_old_path = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R5_2_PH_CodeBeamer\AVE 25.0-R5.2.xlsx'

    f_new_path = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R5_2_PH_CodeBeamer\LBS 22.0.xlsx'

    out_file = r'C:\opt\diff_out.xlsx'

    index_col = 'ID'
    tracked_cols = ['Description', 'Feature']  # Columns to track changed requrements
    # tracked_cols = ['Feature']  # Columns to track changed requrements


    d = Delta2(f_old_path, f_new_path, out_file, index_col, tracked_cols, filter_info_heading=True)
    d.process()

    # Open output file
    excel = 'C:\\Program Files (x86)\\Microsoft Office\\Office16\\EXCEL.EXE'
    subprocess.call(('cmd', '/c', 'start', '', excel, '', out_file))