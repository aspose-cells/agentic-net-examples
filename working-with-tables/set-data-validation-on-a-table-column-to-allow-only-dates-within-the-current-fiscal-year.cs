using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // For ListObject
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsExamples
{
    class FiscalYearDateValidation
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with a header and some dates
                sheet.Cells["A1"].PutValue("TransactionDate");
                sheet.Cells["A2"].PutValue(new DateTime(DateTime.Now.Year, 2, 15));
                sheet.Cells["A3"].PutValue(new DateTime(DateTime.Now.Year, 6, 30));
                sheet.Cells["A4"].PutValue(new DateTime(DateTime.Now.Year, 11, 5));
                sheet.Cells["A5"].PutValue(new DateTime(DateTime.Now.Year - 1, 12, 31)); // outside fiscal year

                // Define the range that will become a table (ListObject)
                int totalRows = 5; // header + 4 data rows
                string tableRange = $"A1:A{totalRows}";

                // Add ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(tableRange, "Transactions", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Determine the current fiscal year start and end dates (assumed calendar year)
                int currentYear = DateTime.Now.Year;
                DateTime fiscalStart = new DateTime(currentYear, 1, 1);
                DateTime fiscalEnd   = new DateTime(currentYear, 12, 31);

                // Get the data range of the table (excluding the header)
                AsposeRange dataRange = table.DataRange; // entire data body of the table
                int startRow = dataRange.FirstRow;
                int endRow   = dataRange.RowCount + startRow - 1;
                int startCol = dataRange.FirstColumn;
                int endCol   = startCol; // single column

                // Create a CellArea covering the column data range
                CellArea validationArea = CellArea.CreateCellArea(startRow, startCol, endRow, endCol);

                // Add a data validation to the worksheet for the defined area
                int validationIndex = sheet.Validations.Add(validationArea);
                Validation validation = sheet.Validations[validationIndex];

                // Configure the validation to allow only dates within the fiscal year
                validation.Type = ValidationType.Date;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = fiscalStart.ToString("yyyy-MM-dd");
                validation.Formula2 = fiscalEnd.ToString("yyyy-MM-dd");
                validation.InputTitle = "Fiscal Year Date";
                validation.InputMessage = $"Enter a date between {fiscalStart:yyyy-MM-dd} and {fiscalEnd:yyyy-MM-dd}.";
                validation.ErrorTitle = "Invalid Date";
                validation.ErrorMessage = "The date is outside the current fiscal year.";
                validation.ShowInput = true;
                validation.ShowError = true;

                // Ensure output directory exists
                string outputPath = "FiscalYearDateValidation.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log exception details for troubleshooting
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}