# -*- coding:utf-8 -*-

'''
Name of Code: PyFISCHERPLOT. Developer: Zongyang Chen, Daming Yang. E-mail: damingyang@sohu.com. Open-source license: Apache License 2.0, Software Require: Python 3.0 language pack, Python library Xlrd, Python library Xlsxwriter.Program Language: Python 3.0. Program Size: 4.48 KB. Purpose: constructing Fischer plots using geological data.
edited by Aqillah Abdul Rahman, Azilah Jahari 
'''

import xlrd
import xlsxwriter
import os
import matplotlib.pyplot as plt
import numpy as np

class WorkBook(object):

    def get_data(self, path):
        print(f"Reading data from file: {path}")
        orign_path = os.path.join(os.path.split(os.path.realpath(__file__))[0], path)
        book = xlrd.open_workbook(orign_path)
        Data_sheet = book.sheets()[0]
        cols_0 = Data_sheet.col_values(0)
        del cols_0[0]
        print(f"Data read successfully from file: {path}")
        return cols_0

    def process_data(self, *data):
        print("Processing data...")
        cols_0 = data[0]
        cols_0 = [round(float(i), 20) for i in cols_0]  # Ensure all values are float and rounded to 20 decimal places
        deep_ave = round(sum(cols_0) / len(cols_0), 20)
        print(f"Mean cycle thickness calculated: {deep_ave}")

        cols_6 = [round(cols_0[i] - deep_ave, 20) for i in range(len(cols_0))]
        cols_7 = []
        for i in range(len(cols_6)):
            if i == 0:
                cols_7.append(cols_6[0])
            else:
                a = round(cols_7[i - 1] + cols_6[i], 20)
                cols_7.append(a)
        cols_8 = [i for i in range(1, len(cols_7) + 1)]
        print("Data processing complete.")
        return cols_0, cols_7, cols_8

    def plot_first_graph(self, cycle_thicknesses):
        print("Plotting first graph...")
        # Compute the cumulative departure from mean cycle thickness (CDMT)
        mean_cycle_thickness = np.mean(cycle_thicknesses)
        cdmt = np.cumsum(cycle_thicknesses - mean_cycle_thickness)

        # Create the plot
        plt.figure(figsize=(10, 6))

        # Plot the thick line for CDMT
        plt.plot(np.cumsum(cycle_thicknesses), cdmt, 'k-', linewidth=2)

        # Plot the triangles for cycle thicknesses below the main line
        for i in range(1, len(cycle_thicknesses)):
            x = [np.cumsum(cycle_thicknesses)[i], np.cumsum(cycle_thicknesses)[i], np.cumsum(cycle_thicknesses)[i-1]]
            y = [cdmt[i], cdmt[i] - abs(cycle_thicknesses[i]), cdmt[i-1]]
            plt.plot(x, y, 'k-', linewidth=1)

        # Add axis labels and title
        plt.xlabel('Thickness (m)')
        plt.ylabel('Cumulative departure from mean cycle thickness (m)')
        plt.title('Thickness vs CDMT')

        # Set axis decimal places
        plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.2f}'))
        plt.gca().xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.2f}'))

        # Show the plot
        plt.grid()
        img_path1 = os.path.join(os.path.split(os.path.realpath(__file__))[0], "plot1.png")
        plt.savefig(img_path1)
        plt.close()
        print("First graph plotted and saved.")
        return img_path1

    def plot_second_graph(self, cols_8, cols_7, cols_0):
        print("Plotting second graph...")
        plt.figure(figsize=(10, 6))
        plt.plot(cols_8, cols_7, 'k-', linewidth=2)
        for i in range(1, len(cols_8)):
            x = [cols_8[i], cols_8[i], cols_8[i - 1]]
            y = [cols_7[i], cols_7[i] - abs(cols_0[i]), cols_7[i - 1]]
            plt.plot(x, y, 'k-', linewidth=1)
        plt.xlabel('Cycle Number')
        plt.ylabel('Cumulative departure from mean cycle thickness (m)')
        plt.title('Cycle Number vs CDMT')

        # Set Y-axis decimal places
        plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.2f}'))

        plt.grid()
        img_path2 = os.path.join(os.path.split(os.path.realpath(__file__))[0], "plot2.png")
        plt.savefig(img_path2)
        plt.close()
        print("Second graph plotted and saved.")
        return img_path2

    def create_excel_with_plots(self, path, *data):
        print(f"Creating Excel file with plots: {path}")
        path_new = os.path.join(os.path.split(os.path.realpath(__file__))[0], "processed" + path)
        cols_0 = data[0][0]
        cols_7 = data[0][1]
        cols_8 = data[0][2]

        workbook_new = xlsxwriter.Workbook(path_new)
        
        # Sheet 1: Data
        worksheet1 = workbook_new.add_worksheet("Data")
        bold = workbook_new.add_format({'bold': 1})
        headings_1 = ["Thickness", "CDMT"]
        data_all = [cols_0, cols_7, cols_8]
        worksheet1.write_row("A1", headings_1, bold)
        worksheet1.write_column("A2", data_all[0])
        worksheet1.write_column("B2", data_all[1])
        headings_2 = ["Cycle Number", "CDMT"]
        worksheet1.write_row("C1", headings_2, bold)
        worksheet1.write_column("C2", data_all[2])
        worksheet1.write_column("D2", data_all[1])
        
        # Sheet 2: First Graph
        worksheet2 = workbook_new.add_worksheet("Thickness vs CDMT")
        img_path1 = self.plot_first_graph(cols_0)
        worksheet2.insert_image('A1', img_path1)

        # Sheet 3: Second Graph
        worksheet3 = workbook_new.add_worksheet("Cycle Number vs CDMT")
        img_path2 = self.plot_second_graph(cols_8, cols_7, cols_0)
        worksheet3.insert_image('A1', img_path2)

        workbook_new.close()
        
        # Remove images after closing the workbook
        os.remove(img_path1)
        os.remove(img_path2)
        
        print("***", path_new, "--Done--", "*****")

if __name__ == '__main__':
    workbook = WorkBook()
    subDirNameList = os.path.split(os.path.realpath(__file__))[0]
    file_paths = os.listdir(subDirNameList)
    for i in file_paths:
        if ".xls" in i:
            try:
                print(f"Processing file: {i}")
                get_data = workbook.get_data(i)
                process_data = workbook.process_data(get_data)
                workbook.create_excel_with_plots(i, process_data)
            except Exception as e:
                print(f'!!!!! {i} --The caught exception is {e}-- !!!!')