import json
import os
import typing
import pandas as pd


class DictList:
    def __init__(self, list_of_dicts: list[dict] = None, logger=None, index_column_name: str = None, root_folder=None):
        self.data_list = list_of_dicts or []
        self.logger = logger or print

        self.index_map = {}  # Dictionary for fast lookups
        self.index_column_name = index_column_name
        if index_column_name:
            self._build_index(index_column_name)

    def __getitem__(self, index):
        return self.data_list[index]

    def __len__(self):
        return len(self.data_list)

    def __iter__(self):
        return iter(self.data_list)

    def add(self, dict_to_add):
        self.data_list.append(dict_to_add)

        # Update index map for new item
        if self.index_column_name and self.index_column_name in dict_to_add:
            self.index_map[dict_to_add[self.index_column_name]] = dict_to_add  # O(1) update

    def get_all(self):
        return self.data_list


    def add_dict_list(self, dict_list):
        """Adds a list of dictionaries and updates the index map."""
        self.data_list.extend(dict_list.data_list)

        # Update index map for new list of items
        if self.index_column_name:
            for item in dict_list.data_list:
                if self.index_column_name in item:
                    self.index_map[item[self.index_column_name]] = item  # O(1) update

        self.logger(f"Added {len(dict_list.data_list)} items to the list.")

    def add_list_of_lists(self, header_list, rows_list):
        """Adds a list of lists and updates the index map."""
        for row in rows_list:
            row_dict = {header: row[idx] for idx, header in enumerate(header_list)}
            self.data_list.append(row_dict)

            # Update index map for new items
            if self.index_column_name and self.index_column_name in row_dict:
                self.index_map[row_dict[self.index_column_name]] = row_dict  # O(1) update

        self.logger(f"Added {len(rows_list)} rows to the list.")

    def remove(self, index):
        if index < 0 or index >= len(self.data_list):
            self.logger(f"Invalid index {index}. No item removed.")
            return

        poped_dict = self.data_list.pop(index)

        # Remove an item from index_map if key exists
        if self.index_column_name:
            contract_code = poped_dict.get(self.index_column_name)
            if contract_code in self.index_map:
                del self.index_map[contract_code]  # O(1) removal

        self.logger(f"Dictionary at index {index} removed successfully.")

    def get_dataframe(self, column_order=None):
        df = pd.DataFrame(self.data_list)
        if column_order:
            for col in column_order:
                if col not in df.columns:
                    df[col] = None
                    self.logger(f"Column '{col}' added with None values.")
            df = df[column_order]
        return df

    def _build_index(self, column_name):
        """Creates a dictionary for quick lookups using 'contract_exchange_code' as the key."""
        self.index_map = {item[column_name]: item for item in self.data_list if column_name in item}

    def find_dict_by_key(self, key, value):
        """Fast lookup using precomputed dictionary if key is 'contract_exchange_code'."""
        if key == self.index_column_name:
            return self.index_map.get(value, None)  # O(1) lookup
        return next((item for item in self.data_list if item.get(key) == value), None)  # O(n) fallback

    def remove_key_from_dict(self, key):
        try:
            df = self.get_dataframe()  # Convert list to DataFrame

            if key in df.columns:
                df = df.drop(columns=[key])

            self.data_list = df.to_dict(orient="records")  # Convert DataFrame back to list

            self.logger(f"Key '{key}' removed successfully from list.")
        except Exception as e:
            self.logger(f"Error while removing key '{key}' from list: {e}")

    def save_as_json(self, file_path, column_order=None):
        try:
            with open(file_path, 'w') as f:
                json.dump(self.data_list, f)
            self.logger(f"Data saved to {file_path}")
        except Exception as e:
            self.logger(f"Error while saving data to {file_path}: {e}")

    def save_as_csv(self, file_path, column_order=None, separator=','):
        try:
            df = self.get_dataframe()
            if column_order:
                df = self.format_column_order(df, column_order)
            df.to_csv(file_path, index=False, sep=separator)
            self.logger(f"Data saved to {file_path}")
            return file_path
        except Exception as e:
            self.logger(f"Error while saving data to {file_path}: {e}")

    def format_column_order(self, df, column_order):
        for col in column_order:
            if col not in df.columns:
                df[col] = None  # Add the missing column with None values
                self.logger(f"Column '{col}' added with None values.")
        df = df[column_order]
        return df

    def save_as_xlsx(self, file_path: str, column_order=None):
        try:
            df = self.get_dataframe(column_order=column_order)

            df.to_excel(file_path, index=False)
            self.logger(f"Data saved to {file_path}")
            return file_path
        except Exception as e:
            self.logger(f"Error while saving data to {file_path}: {e}")

    def save_as_txt(self, file_path: str, column_order=None):
        return self.save_as_csv(file_path, separator='\t', column_order=column_order)

    def save_as_file(self, file_path: str, column_order=None) -> str:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        file_extension = file_path.split('.')[-1]
        self.logger(f"Saving data to {file_path} as {file_extension.title()}.")

        switcher = {
            'xlsx': self.save_as_xlsx,
            'txt': self.save_as_txt,
            'csv': self.save_as_csv,
            'json': self.save_as_json
        }
        try:
            return switcher[file_extension](file_path, column_order=column_order)
        except KeyError:
            self.logger(f"Unsupported file extension: {file_extension}. Supported extensions: {switcher.keys()}")
        except Exception as e:
            self.logger(f"Unexpected error while saving data to {file_path}: {e}")

    def read_dict_list_from_xlsx(self, file_path: str, sheet_name: str = None):
        df = pd.read_excel(file_path, sheet_name=sheet_name) if sheet_name else pd.read_excel(file_path)
        df = df.dropna(how='all')
        self.data_list = df.to_dict(orient="records")

        # Update index map for new items
        if self.index_column_name:
            self._build_index(self.index_column_name)

    def read_dict_list_from_csv(self, file_path: str, sheet_name: str = None):
        df = pd.read_csv(file_path) if sheet_name else pd.read_csv(file_path)
        df = df.dropna(how='all')
        self.data_list = df.to_dict(orient="records")

        # Update index map for new items
        if self.index_column_name:
            self._build_index(self.index_column_name)

    def format_column_names(self, formatter: typing.Callable[[str], str]):
        df = self.get_dataframe()
        df.rename(columns={col: formatter(col) for col in df.columns}, inplace=True)
        self.data_list = df.to_dict(orient="records")

    def sort_by_column(self, key, reverse=False):
        df = self.get_dataframe()
        df.sort_values(by=key, ascending=not reverse, inplace=True)
        self.data_list = df.to_dict(orient="records")

    def close(self):
        self.data_list.clear()
        self.data_list = []
