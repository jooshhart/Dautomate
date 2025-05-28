import pandas as pd

class DataProcessor:
    def __init__(self):
        self.df = None

    def load_file(self, path):
        self.df = pd.read_excel(path)

    def test(self):
        if self.df is not None:
            # Example logic: double the first column if it's numeric
            try:
                col = self.df.columns[0]
                self.df['Processed'] = self.df[col].apply(lambda x: x * 2 if isinstance(x, (int, float)) else x)
            except Exception as e:
                print(f"Error processing data: {e}")

    def save_file(self, path):
        if self.df is not None:
            self.df.to_excel(path, index=False)