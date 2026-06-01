using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Alias to avoid conflict with System.Range
    using AsposeRange = Aspose.Cells.Range;

    public class ApplyCustomNumberFormatToNamedRange
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate financial data in column A, rows 2-5
            sheet.Cells["A2"].PutValue(1234.56);
            sheet.Cells["A3"].PutValue(7890.12);
            sheet.Cells["A4"].PutValue(345.67);
            sheet.Cells["A5"].PutValue(8901.23);

            // Create a named range that refers to the financial data cells
            int nameIndex = workbook.Worksheets.Names.Add("FinancialData");
            Name financialName = workbook.Worksheets.Names[nameIndex];
            financialName.RefersTo = "=Sheet1!$A$2:$A$5";

            // Retrieve the actual Range object from the named range
            AsposeRange financialRange = financialName.GetRange();

            // Create a custom number format style (Euro accounting format as an example)
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "_-€ * #,##0.00_-;_-€ * -#,##0.00_-;_-€ * \"-\"??_-;_-@_-";

            // Apply only the number format part of the style
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            financialRange.ApplyStyle(customStyle, flag);

            // Save the workbook to a file
            string filePath = "FinancialDataFormatted.xlsx";
            workbook.Save(filePath);

            // Verify that the custom format was persisted
            if (File.Exists(filePath))
            {
                Workbook verifyWorkbook = new Workbook(filePath);
                Worksheet verifySheet = verifyWorkbook.Worksheets[0];

                Console.WriteLine("Verifying custom number format applied to named range 'FinancialData':");
                foreach (Cell cell in financialRange)
                {
                    // Get the style of the cell after reloading
                    Style cellStyle = verifySheet.Cells[cell.Name].GetStyle();
                    Console.WriteLine($"{cell.Name}: Value = {cell.Value}, Custom Format = {cellStyle.Custom}");
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
    }
}