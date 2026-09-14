// Title: Copy only cell values from source rows to another range using PasteOptions in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that copies a block of rows to a new location using PasteOptions with PasteType.Values to exclude formulas and formatting. | Show how to use Aspose.Cells PasteOptions to transfer only the values of a source range to a destination range in a worksheet. | Adapt the example to copy rows from one workbook into another workbook while preserving only cell values using PasteOptions.Values.
// Common Searches: Aspose.Cells C# copy rows values only without formulas | How to use PasteOptions.Values to copy a range in Aspose.Cells .NET | Copy rows to another range ignoring formatting Aspose.Cells example | Transfer only cell values between worksheets using Aspose.Cells PasteOptions
// Tags: pasteoptions values copy rows Aspose.Cells | copy rows values only C# Aspose.Cells | ignore formulas range copy .NET | values-only paste type Aspose.Cells | copy range without formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, populates cells with values and formulas, defines source and destination ranges, and uses PasteOptions with PasteType.Values to copy only the cell values—omitting formulas and formatting—from the source rows to the destination rows, then saves the workbook.
    class CopyRowsValuesOnly
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source rows with values and formulas
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["B1"].Formula = "=A1*2"; // Formula that will be ignored
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["B2"].Formula = "=A2*2";
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["B3"].Formula = "=A3*2";

                // Define the source range (rows 0‑2, columns A‑B)
                Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange(0, 0, 3, 2);

                // Define the destination range (starting at row 5, columns A‑B)
                Aspose.Cells.Range destRange = sheet.Cells.CreateRange(5, 0, 3, 2);

                // Configure PasteOptions to copy only values (no formulas, no formatting)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.Values,
                    SkipBlanks = true,
                    OnlyVisibleCells = false,
                    Transpose = false,
                    IgnoreLinksToOriginalFile = true
                };

                // Perform the copy using the specified paste options
                destRange.Copy(sourceRange, pasteOptions);

                // Save the workbook to verify the result
                string outputPath = "CopyRowsValuesOnly.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
