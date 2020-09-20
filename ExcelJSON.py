import re
import pandas as pd
import json


class ExcelJSON(object):
	"""docstring for ExcelJSON"""

	base26 = '123456789abcdefghijklmnopq'
	alphabet = 'abcdefghijklmnopqrstuvwxyz'
	base26map = tuple(zip([i for i in alphabet], [i for i in base26]))

	def __init__(self, location, coordinates=None, sheet=0, index=None, orient="row", rmpadding=True):
		self.location = location
		self.sheet = sheet
		self.index = ExcelJSON._parseCoordinates(self, index)[1]
		self.orient = orient
		self.rmpadding = rmpadding
		self.colfilter = None
		self.rowfilter = None


		if coordinates:
			self.coord = coordinates
			self.coordStart = coordinates.split(":")[0]
			self.coordEnd = coordinates.split(":")[1]
			print(self.coord)
			self.rowStart,self.colStart = ExcelJSON._parseCoordinates(self, self.coordStart)
			self.rowEnd,self.colEnd = ExcelJSON._parseCoordinates(self, self.coordEnd)
		else:
			self.coord = None
			self.rowStart = None 
			self.rowEnd = None 
			self.colStart = None
			self.colEnd = None

	def _parseCoordinates(self, coord):
		"""docstring for parseCoordinates"""
		if coord is None:
			return None, None

		try:
			col = re.search("[a-z]+", coord, flags=re.IGNORECASE)[0]
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

			#convert to base-0 numbering	
			col = convertedCol - 1
		except:
			col = None

		try:
			row = int(re.search("[0-9]+", coord)[0])
			#convert to base-0 numbering
			row += -1
		except:
			row = None	

		# print("Input:", coord)
		# print("Col:",col)
		# print("Row:", row)
		return row, col

	def _dfFormat(self, df):
		"""docstring for parseCoordinates"""

		df = df.where(pd.notnull(df), None)

		# if you're using an index, drop the NaN index rows before you convert to a dictionary
		if self.index is not None:
			df.dropna(subset=[self.index], inplace=True)
			mydict = df.set_index(self.index).to_dict(orient="split")
			output = dict(zip(mydict["index"], [ x for x in mydict["data"]]))
			print("temp dict created: ",output)
			# Trims None values from end of each list. If the list is then empty, delete the dict entry.
			if self.rmpadding == True:
				outputupdate = dict(output)
				for key, value in output.items():
					if len(value) == 1 and type(value[0]) == type(None):
						del(outputupdate[key])
						continue
					elif len(value) == 1:
						outputupdate[key] = value[0]
						continue

					for ele in reversed(value):
						if ele == None and len(value) >= 1:
							del value[-1]
						elif len(value) == 1:
							outputupdate[key] = value[0]
							break
						else:
							outputupdate[key] = value
							break
				output = outputupdate


		# If no index is chosen a list of lists is returned
		else:
			mydict = df.to_dict(orient="split")
			output = [ x for x in mydict["data"]]
			print("temp list created: ", output)
			# Trims None values from end of each list. If the list is then empty, drop the list. If a single value remains, drop the value for the list.
			if self.rmpadding == True:
				outputupdate = list(output)
				for  entry in output:
					if len(entry) == 1 and type(entry[0]) == type(None):
						del(entry)
						continue
					elif len(entry) == 1:
						entry = entry[0]
						continue

					for ele in reversed(entry):
						if ele == None and len(entry) >= 1:
							del entry[-1]
						else:
							break

		return(output)

	def irows(self, myrows):
		rowlist = myrows.split(",")
		rowfinal = []
		for row in rowlist:
			if self.rowStart is not None:
				rowfinal.append(ExcelJSON._parseCoordinates(self, row)[0] - self.rowStart)
			else:
				rowfinal.append(ExcelJSON._parseCoordinates(self, row)[0])
		self.rowfilter = rowfinal

	def icols(self, mycols):
		collist = mycols.split(",")
		colfinal = []
		for col in collist:
			if self.colStart is not None:
				colfinal.append(ExcelJSON._parseCoordinates(self, col)[1] - self.colStart)
			else:
				colfinal.append(ExcelJSON._parseCoordinates(self, col)[1])
		self.colfilter = colfinal 

	def createJSON(self):
		"""docstring for createJSON """
		excel_file = pd.ExcelFile(self.location)
		data = pd.read_excel(excel_file, sheet_name=self.sheet, header=None)

		if self.coord is not None:
			#add 1 to include final row and column in pandas range
			self.rowEnd += 1
			self.colEnd += 1

		#print(data.head(50))
		mydata = data.iloc[self.rowStart:self.rowEnd, self.colStart:self.colEnd]

		if self.rowfilter is not None:
				mydata = mydata.iloc[self.rowfilter,]

		if self.colfilter is not None:
				mydata = mydata.iloc[:,self.colfilter]

		print(mydata)

		dfdict = ExcelJSON._dfFormat(self, mydata)
		formateddict = {"xldata":dfdict}
		finaldata = json.dumps(formateddict, indent=4)
		print("\nfinal data:\n", finaldata)





###Some sample output to play around with
if __name__ == "__main__":
	print("example1:\n-----------------------------\n")
	example1 = ExcelJSON("test.xlsx", sheet="Uniform Data")
	example1.createJSON()

	# TEST DATAFRAM:
	#       0    1    2    3    4    5    6    7    8    9
	# 0    A1   B1   C1   D1   E1   F1   G1   H1   I1   J1
	# 1    A2   B2   C2   D2   E2   F2   G2   H2   I2   J2
	# 2    A3   B3   C3   D3   E3   F3   G3   H3   I3   J3
	# 3    A4   B4   C4   D4   E4   F4   G4   H4   I4   J4
	# 4    A5   B5   C5   D5   E5   F5   G5   H5   I5   J5
	# 5    A6   B6   C6   D6   E6   F6   G6   H6   I6   J6
	# 6    A7   B7   C7   D7   E7   F7   G7   H7   I7   J7
	# 7    A8   B8   C8   D8   E8   F8   G8   H8   I8   J8
	# 8    A9   B9   C9   D9   E9   F9   G9   H9   I9   J9
	# 9   A10  B10  C10  D10  E10  F10  G10  H10  I10  J10
	# 10  A11  B11  C11  D11  E11  F11  G11  H11  I11  J11
	# 11  A12  B12  C12  D12  E12  F12  G12  H12  I12  J12
	# 12  A13  B13  C13  D13  E13  F13  G13  H13  I13  J13
	# 13  A14  B14  C14  D14  E14  F14  G14  H14  I14  J14
	# 14  A15  B15  C15  D15  E15  F15  G15  H15  I15  J15
	# 15  A16  B16  C16  D16  E16  F16  G16  H16  I16  J16
	# 16  A17  B17  C17  D17  E17  F17  G17  H17  I17  J17
	# 17  A18  B18  C18  D18  E18  F18  G18  H18  I18  J18
	# 18  A19  B19  C19  D19  E19  F19  G19  H19  I19  J19
	# 19  A20  B20  C20  D20  E20  F20  G20  H20  I20  J20
	# 20  A21  B21  C21  D21  E21  F21  G21  H21  I21  J21
	# 21  A22  B22  C22  D22  E22  F22  G22  H22  I22  J22
	# 22  A23  B23  C23  D23  E23  F23  G23  H23  I23  J23
	# 23  A24  B24  C24  D24  E24  F24  G24  H24  I24  J24
	# 24  A25  B25  C25  D25  E25  F25  G25  H25  I25  J25
	# 25  A26  B26  C26  D26  E26  F26  G26  H26  I26  J26
	# 26  A27  B27  C27  D27  E27  F27  G27  H27  I27  J27
	# 27  A28  B28  C28  D28  E28  F28  G28  H28  I28  J28
	# 28  A29  B29  C29  D29  E29  F29  G29  H29  I29  J29
	# 29  A30  B30  C30  D30  E30  F30  G30  H30  I30  J30
	#
	# JSON:
	# {'xldata': [['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1'], ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2'], ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3'], ['A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4'], ['A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5', 'J5'], ['A6', 'B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6', 'J6'], ['A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7', 'I7', 'J7'], ['A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8', 'J8'], ['A9', 'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9', 'J9'], ['A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10', 'J10'], ['A11', 'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11', 'I11', 'J11'], ['A12', 'B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12', 'I12', 'J12'], ['A13', 'B13', 'C13', 'D13', 'E13', 'F13', 'G13', 'H13', 'I13', 'J13'], ['A14', 'B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14'], ['A15', 'B15', 'C15', 'D15', 'E15', 'F15', 'G15', 'H15', 'I15', 'J15'], ['A16', 'B16', 'C16', 'D16', 'E16', 'F16', 'G16', 'H16', 'I16', 'J16'], ['A17', 'B17', 'C17', 'D17', 'E17', 'F17', 'G17', 'H17', 'I17', 'J17'], ['A18', 'B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18', 'I18', 'J18'], ['A19', 'B19', 'C19', 'D19', 'E19', 'F19', 'G19', 'H19', 'I19', 'J19'], ['A20', 'B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20', 'I20', 'J20'], ['A21', 'B21', 'C21', 'D21', 'E21', 'F21', 'G21', 'H21', 'I21', 'J21'], ['A22', 'B22', 'C22', 'D22', 'E22', 'F22', 'G22', 'H22', 'I22', 'J22'], ['A23', 'B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23', 'I23', 'J23'], ['A24', 'B24', 'C24', 'D24', 'E24', 'F24', 'G24', 'H24', 'I24', 'J24'], ['A25', 'B25', 'C25', 'D25', 'E25', 'F25', 'G25', 'H25', 'I25', 'J25'], ['A26', 'B26', 'C26', 'D26', 'E26', 'F26', 'G26', 'H26', 'I26', 'J26'], ['A27', 'B27', 'C27', 'D27', 'E27', 'F27', 'G27', 'H27', 'I27', 'J27'], ['A28', 'B28', 'C28', 'D28', 'E28', 'F28', 'G28', 'H28', 'I28', 'J28'], ['A29', 'B29', 'C29', 'D29', 'E29', 'F29', 'G29', 'H29', 'I29', 'J29'], ['A30', 'B30', 'C30', 'D30', 'E30', 'F30', 'G30', 'H30', 'I30', 'J30']]}
	#

	print("\n\nexample2:\n-----------------------------\n")
	example2 = ExcelJSON("test.xlsx", "A1:B25",  index="A", sheet="Uniform Data")
	example2.createJSON()

	# FILTERED DATAFRAM (By range)
	#       0    1
	# 0    A1   B1
	# 1    A2   B2
	# 2    A3   B3
	# 3    A4   B4
	# 4    A5   B5
	# 5    A6   B6
	# 6    A7   B7
	# 7    A8   B8
	# 8    A9   B9
	# 9   A10  B10
	# 10  A11  B11
	# 11  A12  B12
	# 12  A13  B13
	# 13  A14  B14
	# 14  A15  B15
	# 15  A16  B16
	# 16  A17  B17
	# 17  A18  B18
	# 18  A19  B19
	# 19  A20  B20
	# 20  A21  B21
	# 21  A22  B22
	# 22  A23  B23
	# 23  A24  B24
	# 24  A25  B25
	#
	# JSON:
	# {'xldata': {'A1': 'B1', 'A2': 'B2', 'A3': 'B3', 'A4': 'B4', 'A5': 'B5', 'A6': 'B6', 'A7': 'B7', 'A8': 'B8', 'A9': 'B9', 'A10': 'B10', 'A11': 'B11', 'A12': 'B12', 'A13': 'B13', 'A14': 'B14', 'A15': 'B15', 'A16': 'B16', 'A17': 'B17', 'A18': 'B18', 'A19': 'B19', 'A20': 'B20', 'A21': 'B21', 'A22': 'B22', 'A23': 'B23', 'A24': 'B24', 'A25': 'B25'}}


	print("\n\nexample3:\n-----------------------------\n")
	example3 = ExcelJSON("test.xlsx", index="F", sheet="Uniform Data")
	example3.irows("1,2,3,4,5,15,16,17,18,19,20")
	example3.icols("A,B,C,F")
	example3.createJSON()

	#FILTERED DATAFRAM (By using include methods)
	#       0    1    2    5
	# 0    A1   B1   C1   F1
	# 1    A2   B2   C2   F2
	# 2    A3   B3   C3   F3
	# 3    A4   B4   C4   F4
	# 4    A5   B5   C5   F5
	# 14  A15  B15  C15  F15
	# 15  A16  B16  C16  F16
	# 16  A17  B17  C17  F17
	# 17  A18  B18  C18  F18
	# 18  A19  B19  C19  F19
	# 19  A20  B20  C20  F20
	#
	# JSON:
	# {'xldata': {'F1': ['A1', 'B1', 'C1'], 'F2': ['A2', 'B2', 'C2'], 'F3': ['A3', 'B3', 'C3'], 'F4': ['A4', 'B4', 'C4'], 'F5': ['A5', 'B5', 'C5'], 'F15': ['A15', 'B15', 'C15'], 'F16': ['A16', 'B16', 'C16'], 'F17': ['A17', 'B17', 'C17'], 'F18': ['A18', 'B18', 'C18'], 'F19': ['A19', 'B19', 'C19'], 'F20': ['A20', 'B20', 'C20']}}


	print("\n\nexample4:\n-----------------------------\n")
	example4 = ExcelJSON("test.xlsx", index="A")
	#example4.icols("A,C")
	output = example4.createJSON()
	#print(type(output["xldata"]["A3"][3]))

	#FILTERED DATAFRAM (Starting with an excel range, further filtering with include methods)
	#      0    2
	# 0   A1   C1
	# 1   A2   C2
	# 2   A3   C3
	# 3   A4   C4
	# 4   A5   C5
	# 5   A6   C6
	# 6   A7   C7
	# 7   A8   C8
	# 8   A9   C9
	# 9  A10  C10
	#
	# JSON:
	# {'xldata': {'A1': 'C1', 'A2': 'C2', 'A3': 'C3', 'A4': 'C4', 'A5': 'C5', 'A6': 'C6', 'A7': 'C7', 'A8': 'C8', 'A9': 'C9', 'A10': 'C10'}}

	print("\n\nRealworld Example\n______________________________\n")
	x = ExcelJSON(r"test.xlsx", "B2:D421", sheet="Real Example 1", index="B")
	#x.irows("")
	x.icols("B,D")
	x.createJSON()






