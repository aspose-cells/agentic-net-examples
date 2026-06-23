using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsQuotePrefixValidationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define the range where we want to preserve quote prefixes and apply validation (A1:A5)
                string startCell = "A1";
                string endCell = "A5";
                AsposeRange targetRange = cells.CreateRange(startCell, endCell);

                // Populate the range with values that start with a single quote (treated as text)
                for (int row = 0; row < 5; row++)
                {
                    // The leading single quote makes Excel treat the value as text
                    cells[row, 0].PutValue("'12345");
                }

                // Create a style and enable QuotePrefix
                Style quoteStyle = workbook.CreateStyle();
                quoteStyle.QuotePrefix = true; // Mark that the cell value starts with a quote

                // Create a StyleFlag and enable the QuotePrefix flag
                StyleFlag flag = new StyleFlag
                {
                    QuotePrefix = true // Apply only the QuotePrefix property
                };

                // Apply the style to the target range using the flag
                targetRange.ApplyStyle(quoteStyle, flag);

                // Define the cell area for validation (same as targetRange)
                CellArea area = new CellArea
                {
                    StartRow = targetRange.FirstRow,
                    EndRow = targetRange.FirstRow + targetRange.RowCount - 1,
                    StartColumn = targetRange.FirstColumn,
                    EndColumn = targetRange.FirstColumn + targetRange.ColumnCount - 1
                };

                // Add a data validation rule to the same range (whole number between 10 and 20)
                // In some Aspose.Cells versions Add returns the index of the new validation
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.WholeNumber;
                validation.Operator = OperatorType.Between;
                validation.Formula1 = "10";
                validation.Formula2 = "20";

                // Save the workbook
                string outputPath = "QuotePrefixValidationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}