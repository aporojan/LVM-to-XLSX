import xlsxwriter

def extract_data(readfile, writefile):
	buffer = [0 for i in range(columns)]

	with open(readfile, "rt") as f:
		while True:
			line = f.readline().split()
			if len(line) == 0:
				if(len(f.readline().split()) == 0): break

				x.write(",".join([str(i) for i in buffer])+"\n")
				buffer = [0 for i in range(columns)]
				# Time, Hr, Min, Sec, Elapsed, X_Value

			elif(line[0] == "Time"):
				buffer[0] = line[1] # Time
				t = line[1].split(":")
				for i in range(3):
					buffer[i+1] = t[i]

			elif(line[0] == "X_Value"):
				line = f.readline().split()
				buffer[5] = line[0] # X_Value

files = 1 #18
path = "C:/Users/Alex/Downloads/Pressure Data/"

header = "Time,Hour,Min,Sec,Elapsed Time (s),X_Value"
columns = len(header.split(","))
workbook = xlsxwriter.Workbook('Extracted Data.xlsx')


for suffix in range(files):
	with open(path+"Extracted_Data_"+str(suffix)+".csv", "w") as x:
		x.write(header+"\n")
		extract_data(path+"Pressure_Data_"+str(suffix)+".lvm", x)
