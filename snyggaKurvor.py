import matplotlib.pyplot as plt
import openpyxl
import os

def add_x_line(plt, desired_x):
    """
    Adds a vertical dashed line to a given Matplotlib plot.

    Parameters:
    - plt (Matplotlib plot): The plot to which the line will be added.
    - desired_x (float): The x-value at which the vertical dashed line should be drawn.

    Returns:
    - plt (Matplotlib plot): The updated plot with the added line.
    """
    plt.axvline(x=desired_x, color='gray', linestyle='--')
    return plt

def plot_excel_data(plt, 
                    excel_path, 
                    xes_to_highlight, 
                    image_size, 
                    draw_highlight_line = True, 
                    use_grid = True, 
                    savefig = False, 
                    axis_label_font = ('sans-serif', 12), 
                    text_font = ('sans-serif', 10)):
    """
    Plots data from an Excel file using Matplotlib.

    Parameters:
    - plt (Matplotlib plot): The plot to be used for displaying the data.
    - excel_path (str): The path to the Excel file or directory containing Excel files to be plotted.
    - x_to_highlight (float): The x-value at which to highlight data points.
    - image_size (list): The size of the plot image as [width, height].
    - draw_highlight_line (bool): Whether to draw a vertical dashed line at the x_to_highlight value.
    - use_grid (bool): Whether to display a grid on the plot.
    - savefig (bool): Whether to save the plot as an image file.
    - axis_label_font (tuple): Font information for axis labels as (fontname, fontsize).
    - text_font (tuple): Font information for text on the plot as (fontname, fontsize).

    Functionality:
    - Reads data from the specified Excel file(s).
    - Plots the data using Matplotlib, with options for customization.
    - Adds a vertical dashed line at the specified x-value if draw_highlight_line is True.
    - Displays a grid if use_grid is True.
    - Saves the plot as an image file if savefig is True.
    """

    excel_list = []
    excel_filename_list = []

    if ".xlsx" in excel_path:
        excel_list.append(excel_path)
        excel_filename_list.append(excel_path.split("/")[-1])
    else:
        # Iterate over all files in the directory
        for root, dirs, files in os.walk(excel_path):
            for file in files:
                if file.endswith(".xlsx"):
                    excel_list.append(os.path.join(root, file))
                    excel_filename_list.append(file)

    for idx, excel_file in enumerate(excel_list):
        # Specify the name of the sheet where the data is located
        sheet_name = 'Plot Data'

        filename = excel_filename_list[idx]

        # Open the Excel file using openpyxl
        wb = openpyxl.load_workbook(excel_file)
        ws = wb[sheet_name]

        # Find the columns that contain the string "Physical time [s]" in row 4 and use the text to the right as labels
        time_columns = {}
        for col_index, cell in enumerate(ws[4], 1):
            if cell.value == "Physical time [s]":
                time_columns[col_index] = ws.cell(row=4, column=col_index + 1).value

        # Extract the data from the selected time columns and their corresponding data columns
        column_data = []
        column_times = []
        column_labels = []

        for time_column, label in time_columns.items():
            column_times.append([cell[0].value for cell in ws.iter_rows(min_row=12, max_row=ws.max_row, min_col=time_column, max_col=time_column)])
            data_column = time_column + 1
            column_data.append([cell[0].value for cell in ws.iter_rows(min_row=12, max_row=ws.max_row, min_col=data_column, max_col=data_column)])
            column_labels.append(label)

        # Create a line graph using matplotlib
        plt.figure(figsize=(image_size[0], image_size[1]))

        # Plot the data from the selected columns and use the labels from the right as the legend
        for i in range(len(column_data)):
            x = column_times[i]
            y = column_data[i]
            plt.plot(x, y, label=column_labels[i])

            last_value = y[-1]
            color = plt.gca().get_lines()[-1].get_color()
            plt.text(x[-1], last_value, str(int(last_value)), color=color, va='bottom', fontname=text_font[0], fontsize=text_font[1])

        # Set the chart title and axis labels
        plt.title(filename)
        plt.xlabel('Time [s]', fontname=axis_label_font[0], fontsize=axis_label_font[1])
        plt.ylabel('Temperature [°C]', fontname=axis_label_font[0], fontsize=axis_label_font[1])

        # Show the graph with the legend
        plt.legend()
        plt.legend(loc='upper left') # Maybe want it somewhere else? However, it may overlap with the averages


        for x_to_highlight in xes_to_highlight:
            text_count = 0
            # Print the y-values for all lines at the desired x-value
            for i, x_value in enumerate(column_times[0]):  # Using the first column for x-values
                if x_value == x_to_highlight:
                    for j, column_label in enumerate(column_labels):
                        y_value = column_data[j][i]

                        text_x = x[-1] - 50 # Adjust horizontal offset
                        text_y = 15 + (text_count * 8)  # Adjust vertical offset

                        color = plt.gca().get_lines()[j].get_color()
                        plt.text(x_to_highlight + 2, text_y, f'x={x_to_highlight:.3g}, y={y_value:.3g}', color=color, va='bottom', fontname=text_font[0], fontsize=text_font[1])
                        text_count += 1

            # Only draw the line if the user wants it
            if draw_highlight_line:
                plt = add_x_line(plt, x_to_highlight)

        if use_grid:
            plt.grid()

        if savefig:
            path = excel_file.split(".xlsx")[0]
            plt.savefig(f'{path}.png', transparent=True)

        plt.show()

plot_excel_data(plt,
                'C:/Users/avalonuser/Desktop/filer',
                xes_to_highlight = [200, 52],
                image_size = [12, 6],
                draw_highlight_line = True,
                use_grid = True,
                savefig = True,
                axis_label_font = ('montserrat', 10),
                text_font = ('montserrat', 10))
