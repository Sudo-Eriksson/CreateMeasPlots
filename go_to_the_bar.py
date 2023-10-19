import openpyxl
import matplotlib.pyplot as plt
import numpy as np


# Öppna Excel-filen
excel_file = openpyxl.load_workbook('C:/Users/avalonuser/Desktop/dummy_stapel.xlsx')
sheet = excel_file.active

# Skapa tomma listor för data
namn = []
min_varde = []
medelvarde = []
max_varde = []

# Loopa igenom rader i Excel-arket och hämta data
for row in sheet.iter_rows(min_row=2, values_only=True):  # Börja från rad 2 för att undvika rubrikerna
    namn.append(row[0])
    min_varde.append(row[1])
    medelvarde.append(row[2])
    max_varde.append(row[3])

# Antalet namn
antal_namn = len(namn)

# Bredden på staplarna och mellanrummet mellan staplarna
bredd = 0.2
mellanrum = 0.2

# Skapa x-koordinater för staplarna med mellanrum
x = np.arange(antal_namn)

# Skapa ett diagram med tre staplar per namn och ange färger
plt.bar(x - mellanrum, min_varde, width=bredd, label='Min', color='darkslategrey')
plt.bar(x, medelvarde, width=bredd, label='Medel', color='goldenrod')
plt.bar(x + mellanrum, max_varde, width=bredd, label='Max', color='firebrick')

# Lägg till horisontell grid
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Justera x-axeln
plt.xticks(x, namn)

# Lägg till etiketter och en legend
plt.xlabel('Runs')
plt.ylabel('Temperature [°C]')
plt.title("Titel?")
plt.legend()

# Lägg till värdena ovanför staplarna med formatet .3g och samma färger som staplarna
for i in range(antal_namn):
    plt.text(x[i] - mellanrum, min_varde[i], f"{min_varde[i]:.3g}", ha='center', va='bottom', color='darkslategrey')
    plt.text(x[i], medelvarde[i], f"{medelvarde[i]:.3g}", ha='center', va='bottom', color='goldenrod')
    plt.text(x[i] + mellanrum, max_varde[i], f"{max_varde[i]:.3g}", ha='center', va='bottom', color='firebrick')

# Visa diagrammet
plt.show()