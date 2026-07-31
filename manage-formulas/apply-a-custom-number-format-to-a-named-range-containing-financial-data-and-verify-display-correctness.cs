// Title: Apply a Euro Currency Custom Number Format to a Named Range with Aspose.Cells for .NET
// Description: Creates a workbook, fills cells A2:A5 with financial values, defines a named range "FinancialData", builds a Euro currency custom number format, applies only the number‑format part to the range using StyleFlag, prints each cell's displayed string for verification, and saves the file as FinancialDataFormatted.xlsx.
// Keywords: Aspose.Cells | C# custom number format | named range formatting | Euro currency style | StyleFlag number format | verify cell display | .NET spreadsheet API | financial data formatting | apply style to range | StringValue validation
// Common Searches: Aspose.Cells apply custom currency format to named range | How to use StyleFlag for number format only in Aspose.Cells | Verify formatted cell values in C# Aspose.Cells | Create and format named ranges with Euro symbol | Check StringValue after applying custom format Aspose
// Developer Intent: Programmatically set a Euro currency custom number format on a named range and confirm the formatted output before saving.
// Use Cases: Standardize Euro currency display for a column of amounts without altering other cell attributes. | Reuse a single custom format across multiple sheets by applying it to a named range. | Automated testing of number‑format rendering by comparing cell StringValue to expected strings.
// AI Prompts: Generate C# code that creates a named range in Aspose.Cells, applies a USD custom number format, and outputs each cell's formatted string. | Explain how StyleFlag can isolate the number‑format property when applying a style to a range in Aspose.Cells. | Provide a method to programmatically compare a cell's StringValue with an expected formatted value to ensure correctness.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills cells A2:A5 with financial values, defines a named range "FinancialData", builds a Euro currency custom number format, applies only the number‑format part to the range using StyleFlag, prints each cell's displayed string for verification, and saves the file as FinancialDataFormatted.xlsx.
    public class ApplyCustomNumberFormatToNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate financial data (e.g., amounts) in cells A2:A5
                sheet.Cells["A2"].PutValue(1234.56);
                sheet.Cells["A3"].PutValue(7890.12);
                sheet.Cells["A4"].PutValue(-345.67);
                sheet.Cells["A5"].PutValue(0);

                // Create a named range that refers to the financial data cells
                int nameIndex = sheet.Workbook.Worksheets.Names.Add("FinancialData");
                Name financialName = sheet.Workbook.Worksheets.Names[nameIndex];
                financialName.RefersTo = $"={sheet.Name}!$A$2:$A$5";

                // Create a style with a custom number format for currency (Euro)
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "_-€ * #,##0.00_-;_-€ * -#,##0.00_-;_-€ * \"-\"??_-;_-@_-";

                // Use a StyleFlag to apply only the number format part of the style
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Retrieve the range represented by the named range and apply the style
                Aspose.Cells.Range financialRange = financialName.GetRange();
                financialRange.ApplyStyle(customStyle, flag);

                // Verify the applied format by printing the displayed string of each cell
                Console.WriteLine("Verified formatted values in the named range:");
                foreach (Cell cell in financialRange)
                {
                    // StringValue reflects the cell value after applying the number format
                    Console.WriteLine($"{cell.Name}: {cell.StringValue}");
                }

                // Save the workbook to a file
                workbook.Save("FinancialDataFormatted.xlsx");
                Console.WriteLine("Workbook saved as FinancialDataFormatted.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            ApplyCustomNumberFormatToNamedRange.Run();
        }
    }
}
