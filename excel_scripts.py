import pandas as pd
import numpy as np

class Excel_scripts():

    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.df = pd.read_excel(excel_path, header=None, index_col=None)
        self.checked_coordinates = []
        self.table_dict = {}
        self.table_count = 0
        print('Class initialized')

    def get_word_coordinates(self, searched_word):
        for rowno, rows in self.df.iterrows():
            colno = 0
            for value in rows.items():
                if value[1] == searched_word:
                    print(f'Cell {rowno},{colno} contains value {value[1]}')
                    return [rowno, colno]
                colno += 1

    def save_coordinates(self, coordinates):
        if coordinates not in self.checked_coordinates:
            self.checked_coordinates.append(coordinates)

    def find_table(self, start_coordinates):

        if start_coordinates is None:
            start_coordinates = [0,0]

        self.save_coordinates(start_coordinates)

        start_row = start_coordinates[0]
        start_col = start_coordinates[1]

        # print(f'Row numbers of df {df.shape[0]}')
        # print(f'Col numbers of df {df.shape[1]}')
        # print(f'Starting coordinates are {start_row},{start_col}')

        end_row = start_row
        end_col = start_col

        cols_to_left = start_col
        cols_to_right = self.df.shape[1] - start_col - 1
        rows_to_top = start_row
        rows_to_bottom = self.df.shape[0] - start_row - 1

        # print(f'Cells to the left {cols_to_left}')
        # print(f'Cells to the right {cols_to_right}')
        # print(f'Cells to the top {rows_to_top}')
        # print(f'Cells to the bottom {rows_to_bottom}')

        if cols_to_left > 1:
            for i in range(cols_to_left):
                if pd.notna(self.df.iloc[start_row,start_col-1]):
                    start_col = start_col-1
                    # print(f'New start column is {start_col}')
                    self.save_coordinates([start_row, start_col])

        if cols_to_right > 1:
            for i in range(cols_to_right):
                if pd.notna(self.df.iloc[start_row,end_col+1]):
                    end_col = end_col + 1
                    # print(f'New end column is {end_col}')
                    self.save_coordinates([start_row, end_col])

        if rows_to_top > 1:
            for i in range(rows_to_top):
                if pd.notna(self.df.iloc[start_row-1,start_col]):
                    start_row = start_row - 1
                    # print(f'New start row is {start_row}')
                    self.save_coordinates([start_row, start_col])

        if rows_to_bottom > 1:
            for i in range(rows_to_bottom):
                if pd.notna(self.df.iloc[end_row+1,start_col]):
                    end_row = end_row + 1
                    # print(f'New end row is {end_row}')
                    self.save_coordinates([end_row, start_col])

        table = self.df.iloc[start_row:end_row+1, start_col:end_col+1]
        # print(table.shape)

        table = table.reset_index(drop=True)
        table.columns = table.iloc[0]
        table = table[1:]

        return table

    def find_all_tables(self):
        for rowno, rows in self.df.iterrows(): 
            colno = 0
            for value in rows.items():
                if not pd.isna(value[1]):
                    start_coordinates = [rowno, colno]
                    if start_coordinates not in self.checked_coordinates:
                        table = self.find_table(start_coordinates)
                        table_str = table.to_string(index=False)
                        if table.shape[0] >= 2 and table.shape[1] >= 2:
                            if table_str not in [tb[1] for tb in self.table_dict.values()]:
                                self.table_count += 1
                                name = f"table_{self.table_count}"
                                self.table_dict[name] = [table, table_str]
                colno += 1
        print(f'Found {len(self.table_dict.keys())} tables.')

    def get_sheet_count(self):
        try:
            excel_file = pd.ExcelFile(self.excel_path)
            self.sheet_names = excel_file.sheet_names
            return len(self.sheet_names)
        except:
            return None

    def print_all_tables(self):
        for name, table_name in self.table_dict.items():
            print(name)
            print(table_name)