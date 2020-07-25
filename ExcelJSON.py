import re
import numpy as np
import pandas as pd
import json


class ExcelJSON(object):
	"""docstring for ExcelJSON"""

	base26 = '123456789abcdefghijklmnopq'
	alphabet = 'abcdefghijklmnopqrstuvwxyz'
	base26map = tuple(zip([i for i in alphabet], [i for i in base26]))

	def __init__(self, coordinates, location, sheet=0, orient="row"):
		self.location = location
		self.sheet = sheet
		self.orient = orient
		self.coord = coordinates
		self.coordStart = coordinates.split(":")[0]
		self.coordEnd = coordinates.split(":")[1]

	def _parseCoordinates(self, coord):
		"""docstring for parseCoordinates"""
		col = re.search("[a-z]+", coord, flags=re.IGNORECASE)[0]
		row = int(re.search("[0-9]+", coord)[0])

		convertedCol = ""
		for char in col:
			for entry in ExcelJSON.base26map:
				if char.lower() in entry[0]:
					convertedCol += entry[1]
		if convertedCol.count("q") == 1:
			if len(convertedCol) == 1:
				convertedCol = convertedCol.replace("q","p")
				
				convertedCol = int(convertedCol, 26) + 1
			else:
				colBreakout = [i for i in convertedCol]
				if colBreakout.index("q") == 1:
					convertedCol = convertedCol.replace("q","p")
					convertedCol = int(convertedCol, 26) + 26
				else:
					convertedCol = convertedCol.replace("q","p")
					convertedCol = int(convertedCol, 26) + 1
		elif convertedCol.count("q") == 2:
			convertedCol = convertedCol.replace("q","p")
			convertedCol = int(convertedCol, 26) + 26 + 1
		else:
			convertedCol = int(convertedCol, 26)

		print("COORD:", coord)
		print("Col:",convertedCol)
		print("Row:", row)
		row = int(row)
		col = int(convertedCol)
		return row, col



	def createJSON(self, location, rowStart, rowEnd, colStart, colEnd,sheet=0):
		"""docstring for createJSON """
		excel_file = location
		data = pd.read_excel(excel_file, header=None)
		rowStart += -1
		colStart += -1

		print(data.head(20))
		print(type(colEnd))
		mydata = data.iloc[rowStart:rowEnd, colStart:colEnd]
		print(mydata)
		#mydata.set_index("")
		myjson = mydata.to_json(orient="split", index=False)
		print(myjson)
		return(myjson)




x = ExcelJSON("A1:B2", "C:\\Users\\aprayor\\Desktop\\test.xlsx")
rowStart,colStart = x._parseCoordinates("A1")
rowEnd,colEnd = x._parseCoordinates("D5")

output = x.createJSON("C:\\Users\\aprayor\\Desktop\\test.xlsx", rowStart, rowEnd, colStart, colEnd)


