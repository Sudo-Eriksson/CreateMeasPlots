import openpyxl
import matplotlib.pyplot as plt
import numpy as np

def create_bar_chart(file_path, figure_size=(10, 6)):
    """
    Create a bar chart from data in an Excel file.

    Parameters:
    - file_path (str): The path to the Excel file containing the data.
    - figure_size (tuple): Optional. A tuple specifying the size of the figure (width, height) in inches.

    This function reads data from the Excel file and creates a bar chart with three bars (Min, Mean, Max) per data point.
    The chart includes horizontal gridlines and labels. You can customize the figure size using the figure_size parameter.
    """
    # Open the Excel file
    excel_file = openpyxl.load_workbook(file_path)
    sheet = excel_file.active

    # Create empty lists for data
    names = []
    min_value = []
    mean_value = []
    max_value = []

    # Loop through rows in the Excel sheet and retrieve data
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Start from row 2 to skip the headers
        names.append(row[0])
        min_value.append(row[1])
        mean_value.append(row[2])
        max_value.append(row[3])

    # Number of names
    num_names = len(names)

    # Width of the bars and the gap between them
    width = 0.2
    spacing = 0.2

    # Create x-coordinates for the bars with spacing
    x = np.arange(num_names)

    # Set the figure size
    plt.figure(figsize=figure_size)

    # Create a chart with three bars per name and specify colors
    plt.bar(x - spacing, min_value, width=width, label='Min', color='darkslategrey')
    plt.bar(x, mean_value, width=width, label='Mean', color='goldenrod')
    plt.bar(x + spacing, max_value, width=width, label='Max', color='firebrick')

    # Add horizontal gridlines
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    # Adjust the x-axis
    plt.xticks(x, names)

    # Add labels and a legend
    plt.xlabel('Runs')
    plt.ylabel('Temperature [°C]')
    plt.title("Title?")
    plt.legend()

    # Add values above the bars in the format .3g and use the same colors as the bars
    for i in range(num_names):
        plt.text(x[i] - spacing, min_value[i], f"{min_value[i]:.3g}", ha='center', va='bottom', color='darkslategrey')
        plt.text(x[i], mean_value[i], f"{mean_value[i]:.3g}", ha='center', va='bottom', color='goldenrod')
        plt.text(x[i] + spacing, max_value[i], f"{max_value[i]:.3g}", ha='center', va='bottom', color='firebrick')

    # Show the chart
    plt.show()

# Example usage with a custom figure size (e.g., 12x8 inches)
create_bar_chart('C:/Users/avalonuser/Desktop/dummy_stapel.xlsx', figure_size=(10, 6))
