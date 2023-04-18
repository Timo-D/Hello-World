import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import os

fontname = 'Arial'
fontsizeBig = 12
fontsizeSmall = 10
fontsizeTiny = 8

# Excel-Datei einlesen
df = pd.read_excel('C:/Users/Timo/Desktop/Python_Excel/zufallsdaten.xlsx')
dates = df.Datum
calender_weeks = dates
for i, date in enumerate(dates):
    calender_weeks[i] = dates[i].strftime("%W")
    
# Daten nach Kalenderwoche gruppieren
grouped_data = df.groupby('Datum').sum()
grouped_data[1:,]

# Farbpalette erstellen
colors = plt.cm.get_cmap('tab20').colors

# Gestapelte Balken darstellen
fig, ax = plt.subplots(figsize=(250/25.4, 6))
grouped_data.plot(kind='bar', stacked=True, ax=ax, color=colors)


# Achsenbeschriftungen hinzufügen
ax.set_xlabel('Kalenderwoche [KW]', fontsize=fontsizeBig, fontname=fontname, weight ='bold')
ax.set_ylabel('Anzahl der Kundenanforderungen n [-]', fontsize=fontsizeBig, fontname=fontname, weight ='bold')
ax.set_title('Offene Kundenanforderungen', fontsize=fontsizeBig, fontname=fontname, weight ='bold')
# ax.grid(True)
# ax.grid(axis='x')

# Grid zu X- und Y-Achse hinzufügen
ax.grid(axis='y')
ax.set_axisbelow(True)  # Gitter hinter dem Diagramm platzieren
ax.yaxis.set_minor_locator(plt.MultipleLocator(100))  # Minor-Gitter nur für x-Achse
ax.yaxis.set_major_locator(plt.MultipleLocator(200))  # Minor-Gitter nur für x-Achse
#ax.minorticks_on()
ax.grid(which='minor', linestyle=':', linewidth='0.5', color='black')
# ax.grid(which='major', linestyle='-', linewidth='0.5', color='gray')


# Legende hinzufügen
handles, labels = ax.get_legend_handles_labels()
last_values = grouped_data.iloc[-1].values.tolist()

# Letzte Datenpunkte der Spalten in der Legende anzeigen
for i, label in enumerate(labels):
    labels[i] += f" ({last_values[i]})"
    
# Legendenelemente sortieren
sorted_labels = sorted(labels, key=lambda x: (x[1:], x[0])) # Alphanumerisch sortieren
num_rows = int(np.ceil(len(sorted_labels) / 5)) # Anzahl der Reihen berechnen


# Legende in Diagramm einfügen
ax.legend(reversed(handles), sorted_labels, loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=len(labels)/5, prop={'size': fontsizeSmall, 'family': fontname})



# # X-Achse anpassen für KW > 52
# if max(grouped_data.index[2:].astype(int)) > 52:
#     new_ticks = [i for i in range(1, 53)] + [i for i in range(1, max(grouped_data.index)-51)]
#     new_labels = [f"KW {i}\n2022" if i <= 52 else f"KW {i-52}\n2023" for i in new_ticks]
#     ax.set_xticks(range(len(new_ticks)))
#     ax.set_xticklabels(new_labels, fontsize=fontsizeSmall, fontname=fontname)


# Plot anzeigen
plt.show()

# Aktuelles Datum und Uhrzeit abrufen
now = datetime.datetime.now()
date_string = now.strftime("%Y%m%d_%H_%M_%S")


# Aktuelles Arbeitsverzeichnis
current_dir = os.getcwd()

# Absoluter Pfad des Ausgabeordners
output_dir = os.path.join(current_dir, "output")

# Wenn der Ausgabeordner nicht existiert, erstelle ihn
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Diagramm speichern
fig.savefig(os.path.join(output_dir, f"{date_string}_FDB_Test.png"), dpi=300, bbox_inches='tight')