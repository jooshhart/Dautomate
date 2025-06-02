import pandas as pd
import tkinter as tk
import matplotlib.pyplot as plt

class DataProcessor:
    def __init__(self):
        self.df = None

    def load_file(self, path):
        if path.lower().endswith('.csv'):
            self.df = pd.read_csv(path)
        else:
            self.df = pd.read_excel(path)

    def process_efficiency(self):
        if self.df is not None and self.col_map:
            try:
                vin = self.df[self.col_map["vin"]]
                vout = self.df[self.col_map["vout"]]
                iin = self.df[self.col_map["iin"]]
                iout = self.df[self.col_map["iout"]]
                self.df["EFF_CALC (%)"] = (vout * iout) / (vin * iin) * 100
                self.df["PWRL_CALC (mW)"] = (vin * iin - vout * iout)
            except Exception as e:
                print(f"Error calculating efficiency or power loss: {e}")

    def test(self):
        if self.df is not None:
            # Example logic: double the first column if it's numeric
            try:
                col = self.df.columns[0]
                self.df['Processed'] = self.df[col].apply(lambda x: x * 2 if isinstance(x, (int, float)) else x)
            except Exception as e:
                print(f"Error processing data: {e}")

    def save_with_plot(self, path):
        if self.df is not None:
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                self.df.to_excel(writer, index=False, sheet_name='Data')
                workbook = writer.book
                worksheet = writer.sheets['Data']

                # Create the plot and save as image
                fig, ax = plt.subplots()
                ax.scatter(self.df[self.col_map["iout"]], self.df["EFF_CALC (%)"])
                ax.set_xlabel("Load Current (A)")
                ax.set_ylabel("Efficiency (%)")
                ax.set_title("Efficiency Plot")
                plot_img = "efficiency_plot.png"
                fig.savefig(plot_img)
                plt.close(fig)

                # Insert image into Excel
                from openpyxl.drawing.image import Image as XLImage
                img = XLImage(plot_img)
                worksheet.add_image(img, "H2")  # Adjust cell as needed

            # Optionally, remove the temporary image file
            import os
            if os.path.exists(plot_img):
                os.remove(plot_img)