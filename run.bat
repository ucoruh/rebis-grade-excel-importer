@echo off
set /p first_file="CE204.xlsx"
set /p second_file="ExcelNotGirisTaslak4-30-2023.xls"
set /p reference_column="E-Posta"
set /p source_column="Not"
set /p destination_column="Not"
set /p source_sheet=1
set /p destination_sheet=0
rebis-grade-excel-importer.exe %first_file% %second_file% %reference_column% %source_column% %destination_column% %source_sheet% %destination_sheet%
pause