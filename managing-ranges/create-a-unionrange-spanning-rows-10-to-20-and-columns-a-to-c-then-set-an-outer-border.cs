// Title: Aspose.Cells .NET – Create UnionRange A10:C20 and Apply a Thin Black Outline Border
// Description: Demonstrates how to instantiate a Workbook, define a UnionRange that spans rows 10‑20 and columns A‑C on the first worksheet, set a thin black outer border for the range, and save the file as UnionRangeWithBorder.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells UnionRange | C# UnionRange border | outline border Aspose.Cells | A10:C20 range | thin black border .NET | create union range Aspose | Excel border programmatically | Aspose.Cells workbook save
// Common Searches: How to create a UnionRange A10:C20 in Aspose.Cells C# | Set outline border for a UnionRange using Aspose.Cells | Apply thin black border to a specific cell range in .NET | Aspose.Cells example for UnionRange borders | Save workbook after adding range border Aspose
// Developer Intent: Define a UnionRange covering A10:C20 and add a thin black outer border.
// Use Cases: Design a report section with a highlighted block by outlining the data range. | Visually separate a table area in a spreadsheet before exporting to PDF. | Create a printable area with a clear border to improve document layout.
// AI Prompts: Generate C# code to create a UnionRange A10:C20 with a thick red outline border using Aspose.Cells. | Show how to set different border styles for each side of a UnionRange in Aspose.Cells .NET. | Explain how to combine conditional formatting with outline borders on a UnionRange.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsUnionRangeBorderDemo
{
    // Demonstrates how to instantiate a Workbook, define a UnionRange that spans rows 10‑20 and columns A‑C on the first worksheet, set a thin black outer border for the range, and save the file as UnionRangeWithBorder.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Create a UnionRange that spans rows 10‑20 and columns A‑C (address "A10:C20")
            // The second parameter is the worksheet index (0 for the first sheet)
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A10:C20", 0);

            // Apply a thin black outline border to the entire union range
            unionRange.SetOutlineBorders(CellBorderType.Thin, Color.Black);

            // Save the workbook to a file
            workbook.Save("UnionRangeWithBorder.xlsx");
        }
    }
}
