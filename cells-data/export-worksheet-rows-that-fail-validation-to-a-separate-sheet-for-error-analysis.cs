using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsValidationExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and add sample data
                // -------------------------------------------------
                Workbook workbook = new Workbook();                     // create
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "SourceData";

                // Populate some sample data (column A will have whole number validation)
                for (int i = 0; i < 10; i++)
                {
                    srcSheet.Cells[i, 0].PutValue(i * 5);               // valid values: 0,5,10,...
                    srcSheet.Cells[i, 1].PutValue("Row " + i);
                }

                // -------------------------------------------------
                // 2. Add a data validation rule (WholeNumber between 10 and 30) on column A
                // -------------------------------------------------
                // Use the overload that accepts a CellArea (avoids obsolete Add())
                int validationIndex = srcSheet.Validations.Add(new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 9,
                    EndColumn = 0
                });
                Validation validation = srcSheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "10";
                validation.Formula2 = "30";
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Input";
                validation.ErrorMessage = "Value must be between 10 and 30";

                // -------------------------------------------------
                // 3. Identify rows that fail the validation (manual check)
                // -------------------------------------------------
                HashSet<int> errorRows = new HashSet<int>();
                for (int row = 0; row <= 9; row++)
                {
                    // Retrieve the cell value as double (if numeric)
                    object valObj = srcSheet.Cells[row, 0].Value;
                    if (valObj == null) continue;

                    double val;
                    bool isNumber = double.TryParse(valObj.ToString(), out val);
                    if (!isNumber || val < 10 || val > 30)
                    {
                        errorRows.Add(row);
                    }
                }

                // -------------------------------------------------
                // 4. Create a new worksheet to store the error rows
                // -------------------------------------------------
                Worksheet errorSheet = workbook.Worksheets.Add("ErrorRows");

                // Copy header row (if any) – here we copy the first row (index 0) as header
                int destRowIndex = 0;
                srcSheet.Cells.CopyRow(srcSheet.Cells, 0, destRowIndex);
                destRowIndex++;

                // -------------------------------------------------
                // 5. Copy each error row from source to the error sheet
                // -------------------------------------------------
                foreach (int srcRowIndex in errorRows)
                {
                    // Skip the header row if it was already copied
                    if (srcRowIndex == 0) continue;

                    srcSheet.Cells.CopyRow(srcSheet.Cells, srcRowIndex, destRowIndex);
                    destRowIndex++;
                }

                // -------------------------------------------------
                // 6. Save the workbook (lifecycle rule)
                // -------------------------------------------------
                string outputPath = "ValidationErrorExport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}