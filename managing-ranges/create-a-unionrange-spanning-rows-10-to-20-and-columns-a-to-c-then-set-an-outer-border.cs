// Title: Create a UnionRange for rows 10‑20 and columns A‑C and apply a thin black outer border with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create UnionRange for rows 10 to 20 columns A to C | Set outer thin black border on a cell range using Aspose.Cells .NET | Apply outline border to a specific block of cells in Aspose.Cells C# | Define multi‑row multi‑column range with CreateRange in Aspose.Cells | Add border to UnionRange and save workbook as XLSX with Aspose.Cells
// Tags: CreateRange UnionRange Aspose.Cells | outline border cell range .NET | thin black border Aspose.Cells | rows 10-20 columns A-C range Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines a UnionRange that spans rows 10‑20 and columns A‑C, applies a thin black outline border on all four sides of the range, and saves the file as UnionRangeWithBorder.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a range that spans rows 10 to 20 (indices 9 to 19) and columns A to C (indices 0 to 2)
            // Total rows = 20 - 10 + 1 = 11, total columns = 3
            Aspose.Cells.Range unionRange = worksheet.Cells.CreateRange(9, 0, 11, 3);

            // Apply an outer thin black border to the range
            unionRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            unionRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            unionRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            unionRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            // Save the workbook to a file
            workbook.Save("UnionRangeWithBorder.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
