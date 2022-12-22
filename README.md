# LVM-to-XLSX

Simple Python scripts for converting .lvm (LabVIEW measurement) files into .csv (text) or .xlsx (Excel)

**Instructions**:
- change the **files** variable to the number of data files to be converted
- change the **baseName** variable to the common filename, not including numbering
		(ensure data files follow this naming format: _baseName0_, _baseName1_, _baseName2_, etc.)
- change the **path** variable to the folder path of the data files
- run the script!

For example: to convert the files _Extracted_Data_0.lvm_ and _Extracted_Data_1.lvm_ located in _C:/Users/%username%/Downloads_, set **files** = 2, **baseName** = "Extracted_Data_", **path** = "C:/Users/Alex/Downloads"

Planned Features:
- automatically detect number of .lvm files in the same folder as the script and convert all
- include option to save all converted data to the same workbook, separated into worksheets by file (currently forced)
		OR save converted data into separate workbooks (Excel files) matching the number of data files
- develop a simple GUI for running either script with configurable options
