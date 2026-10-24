// Title: C# – Load an Excel workbook with Aspose.Cells, apply a solid fill style to cell A1, and verify the fill pattern
// Common Searches: Aspose.Cells C# set solid background color for a specific cell | How to programmatically check a cell's fill pattern using Aspose.Cells .NET | Apply custom foreground and background colors to an Excel cell with Aspose.Cells | Create a new workbook when file is missing and style a cell in C# | Save an Excel file after modifying cell style with Aspose.Cells
// Tags: Aspose.Cells apply solid fill to Excel cell | C# verify cell background pattern Aspose.Cells | load workbook modify cell style Aspose.Cells | Excel cell foreground background colors C# | Aspose.Cells save workbook after styling

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook if the file is absent), accesses cell A1 on the first worksheet, creates a style with a solid yellow foreground and red background, applies the style, confirms the fill pattern is solid, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet and target cell.
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Create a style with a solid fill (gradient fill is not directly supported).
            Style style = cell.GetStyle();
            style.Pattern = BackgroundType.Solid;          // solid fill
            style.ForegroundColor = Color.Yellow;          // fill color
            // Optionally set background color for contrast (not used in solid fill).
            style.BackgroundColor = Color.Red;

            // Apply the style to the cell.
            cell.SetStyle(style);

            // Verify that the cell's fill is solid.
            Style verifyStyle = cell.GetStyle();
            bool isSolid = verifyStyle.Pattern == BackgroundType.Solid;
            Console.WriteLine("Cell A1 solid fill applied: " + isSolid);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
