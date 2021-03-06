# This file is part of pylabels, a Python library to create PDFs for printing
# labels.
# Copyright (C) 2012, 2013, 2014 Blair Bonnett
#
# pylabels is free software: you can redistribute it and/or modify it under the
# terms of the GNU General Public License as published by the Free Software
# Foundation, either version 3 of the License, or (at your option) any later
# version.
#
# pylabels is distributed in the hope that it will be useful, but WITHOUT ANY
# WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR
# A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License along with
# pylabels.  If not, see <http://www.gnu.org/licenses/>.

import labels
from reportlab.graphics import shapes
import openpyxl

wb = openpyxl.load_workbook('direcciones_solar.xlsx')
ws = wb.get_sheet_by_name('Hoja2')

# Create an A4 portrait (210mm x 297mm) sheets with 2 columns and 8 rows of
# labels. Each label is 90mm x 25mm with a 2mm rounded corner. The margins are
# automatically calculated.
specs = labels.Specification(210, 297, 2, 7, 95, 23, corner_radius=5, top_margin=3, bottom_margin=3)

# Create a function to draw each label. This will be given the ReportLab drawing
# object to draw on, the dimensions (NB. these will be in points, the unit
# ReportLab uses) of the label, and the object to render.
def draw_label(label, width, height, obj):
    # Just convert the object to a string and print this at the bottom left of
    # the label.    
    pos = [50, 40, 30, 20, 10]

    for i in range(5):
        if (obj[i] != None):
            label.add(shapes.String(2, pos[i], str(obj[i]), fontName="Helvetica", fontSize=10))
    # label.add(shapes.String(2, 5, str(obj), fontName="Helvetica", fontSize=10))

# Create the sheet.
sheet = labels.Sheet(specs, draw_label, border=True)

columns = ['B', 'C', 'D', 'E', 'F']

for i in range(79):
    obj = []
    for c in columns:
        obj.append(ws['{0}{1}'.format(c, i+1)].value)

    sheet.add_label(obj)

# Save the file and we are done.
sheet.save('posters_correos.pdf')
print("{0:d} label(s) output on {1:d} page(s).".format(sheet.label_count, sheet.page_count))