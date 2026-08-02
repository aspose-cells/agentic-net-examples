// Title: Aspose.Cells C# – Apply Light Gray Background to Entire Rows of Range C5:C9
// Description: Create a workbook, define the C5:C9 range, expand it to the full rows with EntireRow, style those rows with a solid light‑gray fill, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel range C5:C9 | EntireRow | set row style | background color | light gray fill | CreateRange | SetStyle | Workbook save
// Common Searches: Aspose.Cells set entire row background color C# | apply light gray fill to rows C5 to C9 Aspose.Cells | expand cell range to entire rows Aspose.Cells .NET | C# code for styling rows based on a column range in Excel | Aspose.Cells CreateRange EntireRow example
// Developer Intent: Color the full rows intersecting cells C5‑C9 with light gray using Aspose.Cells.
// Use Cases: Highlight rows 5‑9 across all columns in a generated report. | Add a uniform background to specific rows before exporting to PDF. | Programmatically emphasize rows that meet a business rule in automated Excel output.
// AI Prompts: Generate C# Aspose.Cells code that creates a C5:C9 range, expands it to EntireRow, applies a solid light gray background, and saves the workbook. | Show how to change the fill color or pattern while still targeting the entire rows of a given range in Aspose.Cells. | Explain the steps to use CreateRange and EntireRow to style rows in an Excel file with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsRangeExample
{
    // Create a workbook, define the C5:C9 range, expand it to the full rows with EntireRow, style those rows with a solid light‑gray fill, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet and its cells
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a range that covers cells C5:C9 (column C, rows 5‑9)
                Aspose.Cells.Range columnCRange = cells.CreateRange("C5", "C9");

                // Expand the range to the entire rows intersecting the above range
                Aspose.Cells.Range entireRows = columnCRange.EntireRow;

                // Create a style with a solid light gray background
                Style grayStyle = workbook.CreateStyle();
                grayStyle.Pattern = BackgroundType.Solid;
                grayStyle.ForegroundColor = Color.LightGray;

                // Apply the style to the entire rows range (feature rule: SetStyle)
                entireRows.SetStyle(grayStyle);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("C5_C9_EntireRows_LightGray.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
