import xlsxwriter

def extract_data(readfile, writefile):
	global r
	buffer = [0 for i in range(columns)]	# Initialize local buffer array

	with open(readfile, "rt") as f:
		while True:
			if debug and r == rowsToRead: break	# Code debug condition

				# Read next line
			line = f.readline().split()
			if debug: print(buffer)

			if len(line) == 0:
				# Two empty lines in sequence indicate end of file
				if(len(f.readline().split()) == 0): return

				# Avoid writing zero value first row
				if(buffer[6] == 0): continue

				# Compute elapsed time string s "=E2+60*(C3-C2)+D3-D2" (row 2 example)
				if r > 1: s = "=E{0}+(D{1}-D{0})+60*(C{1}-C{0})+3600*(B{1}-B{0})".format(r, r+1)
				else: s = '0'

				# Flush buffer data into current row
				for c, b in enumerate(buffer):
					if c == 0: writefile.write_string(r, c, b)
					elif c < 4: writefile.write_number(r, c, float(b))
					elif c == 4: writefile.write_formula(r, c, s)
					else: writefile.write_number(r, c, float(b), Scientific_NumFormat)

				# Move to next row and reset buffer
				r += 1
				buffer = [0 for i in range(columns)]
				# Buffer format is [Timestamp, Hr, Min, Sec, Elapsed_Time, X_Value]

			elif(line[0] == "Time"):
				# Absolute Timestamp
				buffer[0] = line[1]
				# Hr, Min, Sec
				buffer[1:4] = line[1].split(":")

			elif(line[0] == "X_Value"):
				# Sample Time & Voltage Measurement (mV)
				buffer[5:7] = [float(i) for i in f.readline().split()]


path = "C:/Users/Alex/Downloads/Lab4_Temperature/"	# Folder path of the data files
baseName = "Lab4_Temperature_"		# Base name structure of data files, not including numbering
files = 13							# Number of data files to be converted

header = "Timestamp,Hour,Min,Sec,Elapsed Time (s),Sample Time (s),Temperature (C)"
columns = len(header.split(","))
rowsToRead = 0	# Code debug condition (0 reads all rows)
debug = False

workbook = xlsxwriter.Workbook(path+'Extracted Data.xlsx')
Time_StrFormat = workbook.add_format()
ThreeDecimal_NumFormat = workbook.add_format().set_num_format('0.000')
Scientific_NumFormat = workbook.add_format().set_num_format('0.00E+00')

for suffix in range(files):
	filename = baseName + str(suffix)
	if debug: print(filename)

	try:
		w = workbook.add_worksheet(str(suffix))
		w.set_column("A:A", None, Time_StrFormat)
		w.set_column("B:E", None, ThreeDecimal_NumFormat)
		w.set_column("F:H", None, Scientific_NumFormat)
		r = 0	# Current row number

		# Write header row
		for c, h in enumerate(header.split(",")):
			w.write(r, c, h)

		r = 1	# Data rows start below the header row
		# Extract relevant data from the .lvm pressure data file
		extract_data(path+filename+".lvm", w)

	except:
		print("Failed to convert data from {}.lvm in {}".format(filename, path))

workbook.close()