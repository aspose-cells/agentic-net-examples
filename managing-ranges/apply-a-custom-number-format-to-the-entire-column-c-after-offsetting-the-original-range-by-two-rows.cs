// Title: Aspose.Cells .NET – Apply Custom Percentage Format to Column C After Row Offset
// Description: Creates a workbook, offsets a range from C1 to C3, retrieves the whole column C, builds a style with the custom format "##0.00%", and applies it using a StyleFlag that targets only the number format.
// Keywords: Aspose.Cells | custom number format | percentage format | column C formatting | range offset rows | StyleFlag | C# Excel automation | .NET workbook styling
// Common Searches: Aspose.Cells offset range by rows | apply custom number format to entire column in .NET | percentage format column C Aspose.Cells | StyleFlag number format only Aspose.Cells | C# set column format after moving range
// Developer Intent: Set a custom percentage number format for the whole column C after moving a source range down two rows.
// Use Cases: Standardize percentage display in financial reports regardless of data start row. | Create a reusable template where column C always shows values as "##0.00%" after inserting rows. | Automate formatting of copied data when the source range is shifted within a worksheet.
// AI Prompts: Generate Aspose.Cells .NET code that offsets a range by two rows and applies the custom format "##0.00%" to the entire column containing the offset range. | Show how to create a Style with a custom percentage pattern and use StyleFlag to apply only the number format to column C after an offset operation. | Write a reusable function that accepts a worksheet, an original range, and a row offset, then formats the whole column of the offset range with a custom numeric pattern.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Alias to avoid conflict with System.Range
    using CellsRange = Aspose.Cells.Range;

    // Creates a workbook, offsets a range from C1 to C3, retrieves the whole column C, builds a style with the custom format "##0.00%", and applies it using a StyleFlag that targets only the number format.
    public class ApplyCustomNumberFormatToColumnC
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define an original range starting at cell C1 (row 0, column 2)
                // For demonstration we use a single‑cell range; the size can be adjusted as needed
                CellsRange originalRange = cells.CreateRange(0, 2, 1, 1);

                // Offset the original range by two rows (row index + 2)
                // Since there is no direct Offset method, create a new range at the offset position
                CellsRange offsetRange = cells.CreateRange(2, 2, 1, 1); // starts at C3

                // Get the entire column that contains the offset range (column C)
                CellsRange entireColumn = offsetRange.EntireColumn;

                // Create a style with a custom number format (percentage with two decimals)
                Style style = workbook.CreateStyle();
                style.Custom = "##0.00%";

                // Configure the style flag to apply only the number format
                StyleFlag styleFlag = new StyleFlag
                {
                    NumberFormat = true
                };

                // Apply the style to the entire column C
                entireColumn.ApplyStyle(style, styleFlag);

                // Save the workbook
                string outputPath = "ColumnC_CustomNumberFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCustomNumberFormatToColumnC.Run();
        }
    }
}
