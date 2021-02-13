"""
Excel et python

Dans ce programme on apprend à faire des graphiques dans la feuille Excel active à partir des données

Éditeur : Laurent REYNAUD
Date : 03-02-21
"""

from openpyxl.workbook import Workbook  # création d'une classeur Excel
from openpyxl import load_workbook  # chargement d'un classeur Excel existant
from openpyxl.chart import PieChart, PieChart3D, Reference, BarChart, BarChart3D, LineChart, LineChart3D  # pour les

# graphiques

"""Chargement du classeur Excel (classeur = workbook = wb)"""
wb = load_workbook('pieces/names.xlsx')

"""Création d'une feuille"""
ws = wb.active

"""Assignation de la fonction pour les différents types diagrammes"""
# chart = PieChart()  # diagramme circulaire
# chart = PieChart3D()   # diagramme circulaire en 3D
# chart = BarChart()  # diagramme à bandes
# chart = BarChart3D()  # diagramme à bandes en 3D
# chart = LineChart()  # diagramme à ligne
chart = LineChart3D()  # diagramme à ligne en 3D

"""Assignation des légendes et des données du graphique"""
labels = Reference(ws, min_col=1, max_col=1, min_row=2, max_row=10)  # légendes à afficher
data = Reference(ws, min_col=3, min_row=1, max_row=10)  # données à représenter en graphique

"""Configuration du graphique"""
chart.add_data(data, titles_from_data=True)  # toujours 'True' sinon les données ne correspondent pas aux légendes
chart.set_categories(labels)

"""Ajout d'un titre"""
chart.title = "Salaires des employés"  # le titre n'apparaît pas ... Calc transformé en Excel ?

"""Ajout et position du graphique dans la feuille Excel"""
ws.add_chart(chart, "E2")  # 'E2' -> placement du coin haut/gauche du graphique

"""Sauvegarde de la feuille"""
wb.save('pieces/names.xlsx')
print('Fichier sauvegardé !')
