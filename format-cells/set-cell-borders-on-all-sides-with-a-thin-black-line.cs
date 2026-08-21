// Title: C# – Add Thin Black Borders to All Sides of a Cell Range with Aspose.Cells
// Description: Creates a workbook, defines range A1:D4, builds a Style, sets thin black borders on left, right, top and bottom, applies only the border settings via StyleFlag, and saves as AllSidesThinBlackBorder.xlsx.
// Keywords: Aspose.Cells C# border | thin black cell border | apply borders to range Aspose.Cells | StyleFlag border only | set cell borders .NET | Excel border formatting code
// Common Searches: Aspose.Cells add thin black border to range | C# set cell borders without changing other styles | How to apply borders on all sides using Aspose.Cells | StyleFlag usage for borders Aspose.Cells .NET | Create thin black borders around A1:D4
// Developer Intent: Add a uniform thin black border to every side of a specified cell range.
// Use Cases: Design a table header where each cell is outlined with a thin black border. | Highlight a financial data block by surrounding it with consistent borders. | Produce a printable invoice section that is clearly separated by a thin black outline.
// AI Prompts: Write C# code with Aspose.Cells to apply a thick red border only to the outer edges of range B2:E10. | Show how to reuse a single Style object to assign different colors to each side of a range in Aspose.Cells. | Explain combining border styling with background fill for a range using Aspose.Cells in .NET.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, defines range A1:D4, builds a Style, sets thin black borders on left, right, top and bottom, applies only the border settings via StyleFlag, and saves as AllSidesThinBlackBorder.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to which borders will be applied
            // Use fully qualified type to avoid conflict with System.Range
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:D4");

            // Create a style object
            Style style = workbook.CreateStyle();

            // Set thin black borders on all four sides
            style.SetBorder(BorderType.LeftBorder,   CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder,  CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder,    CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

            // Apply only the border settings to the range
            StyleFlag flag = new StyleFlag { Borders = true };
            range.ApplyStyle(style, flag);

            // Save the workbook (save rule)
            string outputPath = "AllSidesThinBlackBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
