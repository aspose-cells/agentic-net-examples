// Title: Aspose.Cells for .NET – Apply Thin Black Borders to All Sides of a Cell Range (C#)
// Description: C# example that creates a workbook, defines a range (e.g., A1:D4), builds a Style, sets thin black borders on left, right, top and bottom, uses a StyleFlag to affect only borders, applies the style to the range, and saves the file as AllBordersThinBlack.xlsx.
// Keywords: Aspose.Cells C# border | set cell borders Aspose.Cells | thin black border .NET | apply borders to range | StyleFlag borders | Excel cell formatting C# | Aspose.Cells workbook styling | border all sides Aspose.Cells
// Common Searches: Aspose.Cells add thin black border to range C# | how to set borders on all sides of cells using Aspose.Cells | C# code for applying borders with StyleFlag in Aspose.Cells | apply uniform border to Excel range Aspose.Cells .NET | set thin black cell borders Aspose.Cells example
// Developer Intent: Add a uniform thin black border to every edge of a specified cell range in an Aspose.Cells workbook.
// Use Cases: Create printable tables with clear grid lines for reports. | Emphasize header or summary blocks by framing them with a consistent border. | Prepare spreadsheets for export where each data section needs visual separation.
// AI Prompts: Generate C# code that applies a thick red border only to the top edge of a selected range using Aspose.Cells. | Show how to assign different border styles (thin, medium, dashed) to each side of a range in Aspose.Cells for .NET. | Provide an example of applying borders to multiple non‑contiguous ranges with a single StyleFlag in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// C# example that creates a workbook, defines a range (e.g., A1:D4), builds a Style, sets thin black borders on left, right, top and bottom, uses a StyleFlag to affect only borders, applies the style to the range, and saves the file as AllBordersThinBlack.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to which the borders will be applied (e.g., A1:D4)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:D4");

            // Create a style object
            Style style = workbook.CreateStyle();

            // Set thin black borders on all four sides
            style.SetBorder(BorderType.LeftBorder,   CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder,  CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder,    CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            // Prepare a StyleFlag to indicate that only border settings should be applied
            StyleFlag flag = new StyleFlag();
            flag.Borders = true;

            // Apply the style with borders to the defined range
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("AllBordersThinBlack.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
