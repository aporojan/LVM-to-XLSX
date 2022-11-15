import xlsxwriter

def extract_data(readfile, writefile):
	global r
	buffer = [0 for i in range(columns)]
			# Initialize local buffer array

	with open(readfile, "rt") as f:
		while True:
			if r == rowsToRead: break	# Code debug condition

			# Read next line
			line = f.readline().split()	

			if len(line) == 0:
				# Two empty lines in sequence indicate end of file
				if(len(f.readline().split()) == 0): return

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
				# Buffer format is [Time, Hr, Min, Sec, Elapsed_Time, X_Value]

			# Time data
			elif(line[0] == "Time"):
				buffer[0] = line[1]
				t = line[1].split(":")
				for i in range(3):
					buffer[i+1] = t[i]

			# X0 value
			elif(line[0] == "X0"):
				buffer[5] = line[1]

			# X_Value value
			elif(line[0] == "X_Value"):
				line = f.readline().split()
				buffer[6] = line[1]
				buffer[7] = 1000 * float(line[1])


path = "C:/Users/Alex/Downloads/Pressure Data/"
files = 18		# 18 total
rowsToRead = 0	# Code debug condition (0 reads all rows)

header = "Time,Hour,Min,Sec,Elapsed Time (s),X0,X_Value,Pressure (psi)"
columns = len(header.split(","))

workbook = xlsxwriter.Workbook(path+'Extracted Pressure Data.xlsx')
Time_StrFormat = workbook.add_format()
ThreeDecimal_NumFormat = workbook.add_format().set_num_format('0.000')
Scientific_NumFormat = workbook.add_format().set_num_format('0.00E+00')

for suffix in range(files):

	w = workbook.add_worksheet("Extracted_Data_"+str(suffix))
	w.set_column("A:A", None, Time_StrFormat)
	w.set_column("B:E", None, ThreeDecimal_NumFormat)
	w.set_column("F:H", None, Scientific_NumFormat)
	r = 0	# Current row number

	# Write header row
	for c, h in enumerate(header.split(",")):
		w.write(r, c, h)

	r = 1	# Data rows start below the header row
	extract_data(path+"Pressure_Data_"+str(suffix)+".lvm", w)
	# Extract relevant data from the .lvm pressure data file

workbook.close()