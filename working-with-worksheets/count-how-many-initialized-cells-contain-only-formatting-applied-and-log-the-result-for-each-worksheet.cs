// Title: C# – Count Formatting‑Only Initialized Cells per Worksheet with Aspose.Cells
// Description: Shows how to create or load a workbook, apply custom styles, then enumerate only the instantiated cells of each worksheet to identify cells that have no value but a non‑default style, log the count per sheet, and optionally save the file.
// Keywords: Aspose.Cells C# count formatted only cells | enumerate instantiated cells .NET | detect empty styled cells Aspose | worksheet style‑only cell count | audit workbook formatting Aspose.Cells | reduce Excel file size by clearing styles | C# Excel style comparison default | Aspose.Cells cell style analysis
// Common Searches: how to count cells with only formatting using Aspose.Cells | Aspose.Cells enumerate initialized cells without values | C# find empty cells that have custom style in a workbook | count style‑only cells per worksheet Aspose | Aspose.Cells audit formatting‑only cells
// Developer Intent: Determine the number of initialized cells that contain a custom style but no data in each worksheet of an Excel workbook.
// Use Cases: Validate that a template’s styling does not contain stray formatted cells before distribution. | Generate a styling audit report that lists formatting‑only cell counts per sheet. | Identify and clear empty formatted cells to shrink workbook size and improve performance.
// AI Prompts: Write a method that returns a dictionary of worksheet names and the count of cells that have a custom style but no value using Aspose.Cells. | Adapt the example to ignore hidden worksheets while counting formatting‑only cells. | Provide code that removes the style from cells with no value after the count is logged.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create or load a workbook, apply custom styles, then enumerate only the instantiated cells of each worksheet to identify cells that have no value but a non‑default style, log the count per sheet, and optionally save the file.
    class CountFormattedOnlyCells
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // lifecycle: create

                // Example: add some data and formatting to demonstrate the counting logic
                Worksheet sheet1 = workbook.Worksheets[0];
                Cells cells1 = sheet1.Cells;

                // Cell with only formatting (no value)
                Style fmtOnly = workbook.CreateStyle();
                fmtOnly.Font.IsBold = true;
                cells1["B2"].SetStyle(fmtOnly);

                // Cell with value and formatting
                Style fmtBoth = workbook.CreateStyle();
                fmtBoth.Font.Color = System.Drawing.Color.Blue;
                cells1["C3"].PutValue(123);
                cells1["C3"].SetStyle(fmtBoth);

                // Cell with only value
                cells1["D4"].PutValue("Text");

                // Process each worksheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    long formattedOnlyCount = 0;

                    // Enumerate only instantiated cells
                    foreach (Cell cell in ws.Cells)
                    {
                        // Determine if the cell has no value
                        bool hasNoValue = cell.Value == null || string.IsNullOrEmpty(cell.StringValue);

                        // Determine if the cell has any formatting applied (different from default style)
                        Style cellStyle = cell.GetStyle();
                        bool hasFormatting = !cellStyle.Equals(workbook.DefaultStyle);

                        if (hasNoValue && hasFormatting)
                        {
                            formattedOnlyCount++;
                        }
                    }

                    Console.WriteLine($"Worksheet \"{ws.Name}\": {formattedOnlyCount} initialized cells contain only formatting.");
                }

                // Save the workbook if needed (lifecycle: save)
                string outputPath = "FormattedOnlyCellsCount.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
