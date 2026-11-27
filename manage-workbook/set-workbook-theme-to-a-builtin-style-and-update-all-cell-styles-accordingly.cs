// Title: Apply a blue font and light‑blue solid fill to a cell and save the workbook as XLSX using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set cell font color to blue and background to light blue | How to create and assign a custom style to a single cell with Aspose.Cells | C# example saving an Excel workbook after styling a cell using Aspose.Cells | Apply solid fill pattern to a cell in Aspose.Cells for .NET | Create a new workbook, style a cell, and export to XLSX with Aspose.Cells
// Tags: apply custom cell style Aspose.Cells | set font color and fill Aspose.Cells C# | create solid background style Aspose.Cells | save workbook as XLSX Aspose.Cells | initialize workbook and worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The sample creates a new workbook, writes "Sample Text" to cell A1, defines a style with a blue 12‑point font and a light‑blue solid fill, applies the style to the cell, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample Text");

            // Create a style with custom colors (replace theme colors)
            Style style = workbook.CreateStyle();
            style.Font.Color = Color.Blue;          // font color
            style.Font.Size = 12;
            style.ForegroundColor = Color.LightBlue; // fill color
            style.Pattern = BackgroundType.Solid;
            cell.SetStyle(style);

            // Define output file path
            string outputPath = "Result.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
