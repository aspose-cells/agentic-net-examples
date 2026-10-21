// Title: Apply a custom number format to column C starting from row 3 with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply custom number format to column C after skipping first two rows | CreateRange offset by two rows and set number format in Aspose.Cells .NET | Set Excel column C number format beginning at row 3 using Aspose.Cells API | Apply style with custom number format to a column range in Aspose.Cells C# example | How to format cells in column C from row 3 onward with Aspose.Cells for .NET
// Tags: custom number format style Aspose.Cells | CreateRange with row offset Aspose.Cells | StyleFlag NumberFormat property Aspose.Cells | format column C Excel using Aspose.Cells | skip initial rows apply style Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, calculates the last used row, creates a range that covers column C starting at the third row (offset by two rows), defines a style with the custom number format "#,##0.00", applies the style to the range using a StyleFlag that targets only the number format, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last used row in the sheet
                int lastRow = sheet.Cells.MaxDataRow;

                // Ensure there are rows beyond the offset row (row index 2)
                int startRow = 2; // zero‑based index (third row)
                int totalRows = Math.Max(0, lastRow - startRow + 1);
                if (totalRows == 0)
                {
                    Console.WriteLine("No data rows found after the specified offset.");
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                    return;
                }

                // Create a range that represents column C (index 2) offset by two rows
                Aspose.Cells.Range offsetRange = sheet.Cells.CreateRange(startRow, 2, totalRows, 1);

                // Create a style with a custom number format
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "#,##0.00";

                // Define which style attributes to apply (only number format)
                StyleFlag flag = new StyleFlag
                {
                    NumberFormat = true
                };

                // Apply the custom style to the offset range
                offsetRange.ApplyStyle(customStyle, flag);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
