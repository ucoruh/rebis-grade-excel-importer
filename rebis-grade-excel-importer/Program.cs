using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check if the user provided the correct number of arguments
            if (args.Length != 7)
            {
                Console.WriteLine("Please provide five arguments: first file path, second file path, reference column name, source column name, destination column name, first file sheet index, second file sheet index");
                return;
            }

            // Get the file paths and column names from the command line arguments
            string firstFilePath = args[0];
            string secondFilePath = args[1];
            string referenceColumn = args[2];

            string sourceColumn = args[3];
            string destinationColumn = args[4];
            int firstFileSheetIndex = Convert.ToInt32(args[5]);
            int secondFileSheetIndex = Convert.ToInt32(args[6]);

            // Load the first Excel file into memory
            IWorkbook firstWorkbook = null;
            if (Path.GetExtension(firstFilePath) == ".xls")
            {
                using (FileStream file = new FileStream(firstFilePath, FileMode.Open, FileAccess.Read))
                {
                    firstWorkbook = new HSSFWorkbook(file);
                }
            }
            else if (Path.GetExtension(firstFilePath) == ".xlsx")
            {
                using (FileStream file = new FileStream(firstFilePath, FileMode.Open, FileAccess.Read))
                {
                    firstWorkbook = new XSSFWorkbook(file);
                }
            }
            else
            {
                Console.WriteLine("The first file must be an XLS or XLSX file.");
                return;
            }

            ISheet firstWorksheet = firstWorkbook.GetSheetAt(firstFileSheetIndex);

            // Find the column index of the reference column in the first file
            int referenceColumnIndex = 0;
            IRow firstHeaderRow = firstWorksheet.GetRow(0);
            for (int i = 0; i < firstHeaderRow.LastCellNum; i++)
            {
                ICell cell = firstHeaderRow.GetCell(i);
                if (cell.StringCellValue.ToLower() == referenceColumn.ToLower())
                {
                    referenceColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            // Find the column index of the source column in the first file
            int sourceColumnIndex = 0;
            for (int i = 0; i < firstHeaderRow.LastCellNum; i++)
            {
                ICell cell = firstHeaderRow.GetCell(i);
                if (cell.StringCellValue.ToLower() == sourceColumn.ToLower())
                {
                    sourceColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            // Load the second Excel file into memory
            IWorkbook secondWorkbook = null;
            if (Path.GetExtension(secondFilePath) == ".xls")
            {
                using (FileStream file = new FileStream(secondFilePath, FileMode.Open, FileAccess.Read))
                {
                    secondWorkbook = new HSSFWorkbook(file);
                }
            }
            else if (Path.GetExtension(secondFilePath) == ".xlsx")
            {
                using (FileStream file = new FileStream(secondFilePath, FileMode.Open, FileAccess.Read))
                {
                    secondWorkbook = new XSSFWorkbook(file);
                }
            }
            else
            {
                Console.WriteLine("The second file must be an XLS or XLSX file.");
                return;
            }

            ISheet secondWorksheet = secondWorkbook.GetSheetAt(secondFileSheetIndex);

            // Find the column index of the reference column in the second file
            int secondReferenceColumnIndex = 0;
            IRow secondHeaderRow = secondWorksheet.GetRow(0);
            for (int i = 0; i < secondHeaderRow.LastCellNum; i++)
            {
                ICell cell = secondHeaderRow.GetCell(i);
                if (cell.StringCellValue.ToLower() == referenceColumn.ToLower())
                {
                    secondReferenceColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            // Find the column index of the destination column in the second file
            int destinationColumnIndex = 0;
            for (int i = 0; i < secondHeaderRow.LastCellNum; i++)
            {
                ICell cell = secondHeaderRow.GetCell(i);
                if (cell.StringCellValue.ToLower() == destinationColumn.ToLower())
                {
                    destinationColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            // Loop through each row in the second file and copy the source column value to the destination column if the reference column value matches
            for (int i = 1; i <= secondWorksheet.LastRowNum; i++)
            {
                IRow row = secondWorksheet.GetRow(i);
                ICell referenceCell = row.GetCell(secondReferenceColumnIndex);
                string referenceValue = referenceCell.StringCellValue.Trim();

                // Loop through each row in the first file to find the matching reference value
                for (int j = 1; j <= firstWorksheet.LastRowNum; j++)
                {
                    IRow firstRow = firstWorksheet.GetRow(j);
                    ICell firstReferenceCell = firstRow.GetCell(referenceColumnIndex);
                    string firstReferenceValue = firstReferenceCell.StringCellValue.Trim();

                    // If the reference values match, copy the source column value to the destination column in the second file
                    if (referenceValue.Equals(firstReferenceValue, StringComparison.OrdinalIgnoreCase))
                    {
                        ICell sourceCell = firstRow.GetCell(sourceColumnIndex);
                        ICell destinationCell = row.GetCell(destinationColumnIndex);

                        if (destinationCell.CellType == CellType.Numeric)
                        {
                            // Destination cell already contains a numeric value
                            double numericValue = sourceCell.NumericCellValue;
                            destinationCell.SetCellValue(numericValue);
                        }
                        else if (destinationCell.CellType == CellType.String)
                        {
                            // Destination cell already contains a string value
                            string stringValue = sourceCell.ToString();
                            destinationCell.SetCellValue(stringValue);
                        }
                        else
                        {
                            // Destination cell is empty or contains another data type
                            if (sourceCell.CellType == CellType.Numeric)
                            {
                                // Convert numeric value to double
                                double numericValue = sourceCell.NumericCellValue;
                                destinationCell.SetCellValue(numericValue);
                            }
                            else if (sourceCell.CellType == CellType.String)
                            {
                                // Convert string value to double if possible, otherwise set as string
                                string stringValue = sourceCell.ToString();
                                if (double.TryParse(stringValue, out double numericValue))
                                {
                                    destinationCell.SetCellValue(numericValue);
                                }
                                else
                                {
                                    destinationCell.SetCellValue(stringValue);
                                }
                            }
                        }


                        break;
                    }
                }
            }

            // Save the changes to the second file
            using (FileStream file = new FileStream(secondFilePath, FileMode.Create, FileAccess.Write))
            {
                secondWorkbook.Write(file,false);
            }

            Console.WriteLine("Done!");
        }
    }

}