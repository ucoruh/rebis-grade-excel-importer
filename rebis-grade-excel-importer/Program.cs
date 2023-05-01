using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

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
            int referenceColumnIndex = -1;
            IRow firstHeaderRow = firstWorksheet.GetRow(0);
            for (int i = 0; i < firstHeaderRow.LastCellNum; i++)
            {
                NPOI.SS.UserModel.ICell cell = firstHeaderRow.GetCell(i);

                if (cell.CellType == CellType.Numeric)
                {
                    Console.WriteLine("Numeric Header Detected Row:["+cell.RowIndex+"] Column:["+cell.ColumnIndex+"], Skipping in " + firstFilePath + " file");
                    continue;
                }
                    
                if (cell.StringCellValue.ToLower() == referenceColumn.ToLower())
                {
                    referenceColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            if(referenceColumnIndex== -1)
            {
                Console.WriteLine("Reference Column [" + referenceColumn + "] value not found in "+ firstFilePath + " file.");
                Console.WriteLine("Please Check column names if its not exist create it");
                return;
            }

            // Find the column index of the source column in the first file
            int sourceColumnIndex = -1;
            for (int i = 0; i < firstHeaderRow.LastCellNum; i++)
            {
                NPOI.SS.UserModel.ICell cell = firstHeaderRow.GetCell(i);

                if (cell.CellType == CellType.Numeric)
                {
                    Console.WriteLine("Numeric Header Detected Row:[" + cell.RowIndex + "] Column:[" + cell.ColumnIndex + "], Skipping in " + firstFilePath + " file");
                    continue;
                }

                if (cell.StringCellValue.ToLower() == sourceColumn.ToLower())
                {
                    sourceColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            if (sourceColumnIndex == -1)
            {
                Console.WriteLine("Reference Column [" + sourceColumn + "] value not found in " + firstFilePath + " file.");
                Console.WriteLine("Please Check column names if its not exist create it");
                return;
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
            int secondReferenceColumnIndex = -1;
            IRow secondHeaderRow = secondWorksheet.GetRow(0);
            for (int i = 0; i < secondHeaderRow.LastCellNum; i++)
            {
                NPOI.SS.UserModel.ICell cell = secondHeaderRow.GetCell(i);

                if (cell.CellType == CellType.Numeric)
                {
                    Console.WriteLine("Numeric Header Detected Row:[" + cell.RowIndex + "] Column:[" + cell.ColumnIndex + "], Skipping in " + secondFilePath + " file");
                    continue;
                }

                if (cell.StringCellValue.ToLower() == referenceColumn.ToLower())
                {
                    secondReferenceColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            if (secondReferenceColumnIndex == -1)
            {
                Console.WriteLine("Reference Column [" + referenceColumn + "] value not found in " + secondFilePath + " file.");
                Console.WriteLine("Please Check column names if its not exist create it");
                return;
            }



            // Find the column index of the destination column in the second file
            int destinationColumnIndex = -1;
            for (int i = 0; i < secondHeaderRow.LastCellNum; i++)
            {
                NPOI.SS.UserModel.ICell cell = secondHeaderRow.GetCell(i);

                if (cell.CellType == CellType.Numeric)
                {
                    Console.WriteLine("Numeric Header Detected Row:[" + cell.RowIndex + "] Column:[" + cell.ColumnIndex + "], Skipping in " + secondFilePath + " file");
                    continue;
                }

                if (cell.StringCellValue.ToLower() == destinationColumn.ToLower())
                {
                    destinationColumnIndex = cell.ColumnIndex;
                    break;
                }
            }

            if (destinationColumnIndex == -1)
            {
                Console.WriteLine("Reference Column [" + destinationColumn + "] value not found in " + secondFilePath + " file.");
                Console.WriteLine("Please Check column names if its not exist create it");
                return;
            }

            // Loop through each row in the second file and copy the source column value to the destination column if the reference column value matches
            for (int i = 1; i <= secondWorksheet.LastRowNum; i++)
            {
                IRow row = secondWorksheet.GetRow(i);
                NPOI.SS.UserModel.ICell referenceCell = row.GetCell(secondReferenceColumnIndex);

                if (referenceCell == null)
                {
                    continue;
                }
                
                string referenceValue = referenceCell.StringCellValue.Trim();

                // Loop through each row in the first file to find the matching reference value
                for (int j = 1; j <= firstWorksheet.LastRowNum; j++)
                {
                    IRow firstRow = firstWorksheet.GetRow(j);
                    NPOI.SS.UserModel.ICell firstReferenceCell = firstRow.GetCell(referenceColumnIndex);

                    if(firstReferenceCell==null)
                    {
                        Console.WriteLine( "Cell is Null Row["+j+"]Column["+referenceColumnIndex+"] Skipping in " + firstFilePath + " file.");
                        continue;
                    }

                    string firstReferenceValue = firstReferenceCell.StringCellValue.Trim();

                    // If the reference values match, copy the source column value to the destination column in the second file
                    if (referenceValue.Equals(firstReferenceValue, StringComparison.OrdinalIgnoreCase))
                    {
                        NPOI.SS.UserModel.ICell sourceCell = firstRow.GetCell(sourceColumnIndex);
                        NPOI.SS.UserModel.ICell destinationCell = row.GetCell(destinationColumnIndex);

                        if(sourceCell == null)
                        {
                            Console.WriteLine("Source Cell is Null Row[" + j + "]Column[" + sourceColumnIndex + "] Skipping in " + firstFilePath + " file.");
                            continue;
                        }

                        if(destinationCell == null)
                        {
                            Console.WriteLine("Destination Cell is Null Row[" + j + "]Column[" + destinationColumnIndex + "] Creating in " + secondFilePath + " file.");
                            destinationCell = row.CreateCell(destinationColumnIndex);
                            destinationCell.SetCellValue(0);
                        }

                        if (sourceCell.CellType == CellType.Formula) // check if the cell contains a formula
                        {
                            sourceCell.SetCellType(CellType.Numeric); // set the cell type to Formula

                            try
                            {
                                double numericValue = sourceCell.NumericCellValue;
                                destinationCell.SetCellValue(numericValue);
                            }
                            catch
                            {
                                destinationCell.SetCellValue(0);
                            }

                        }
                        else if (sourceCell.CellType == CellType.Numeric)
                        {
                            // Destination cell already contains a numeric value
                            try
                            {
                                double numericValue = sourceCell.NumericCellValue;
                                destinationCell.SetCellValue(numericValue);
                            }
                            catch
                            {
                                destinationCell.SetCellValue(0);
                            }

                        }
                        else if (destinationCell.CellType == CellType.String)
                        {
                            // Destination cell already contains a string value
                            string stringValue = sourceCell.StringCellValue.ToString();
                            destinationCell.SetCellValue(stringValue);
                        }
                        else
                        {

                            if (sourceCell.CellType == CellType.Numeric)
                            {
                                // Convert numeric value to double
                                try
                                {
                                    double numericValue = sourceCell.NumericCellValue;
                                    destinationCell.SetCellValue(numericValue);
                                }
                                catch
                                {
                                    destinationCell.SetCellValue(0);
                                }

                            }
                            else if (sourceCell.CellType == CellType.String)
                            {
                                // Convert string value to double if possible, otherwise set as string
                                string stringValue = sourceCell.StringCellValue.ToString();
                                if (double.TryParse(stringValue, out double numericValue))
                                {
                                    destinationCell.SetCellValue(numericValue);
                                }
                                else
                                {
                                    destinationCell.SetCellValue(stringValue);
                                }
                            }
                            else
                            {
                                destinationCell.SetCellValue(0);
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

            Console.WriteLine("Operation Completed!");
        }
    }

}