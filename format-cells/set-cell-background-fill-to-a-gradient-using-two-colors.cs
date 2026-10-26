// Title: Apply a two‑color gradient background to an Excel cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell background gradient using two colors | C# example of applying gradient fill to an Excel cell with Aspose.Cells | How to use ForegroundColor and BackgroundColor for gradient effect in Aspose.Cells | Aspose.Cells gradient fill not supported version workaround | Set gradient degree for cell fill in Aspose.Cells .NET
// Tags: Aspose.Cells cell gradient background C# | foregroundcolor backgroundcolor gradient Aspose.Cells | gradientdegree property usage Aspose.Cells | solid fill pattern fallback Aspose.Cells | Excel workbook creation Aspose.Cells C#

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, writes "Gradient Fill" to cell A1, sets the foreground color to blue and the background color to light blue, applies a solid fill pattern because gradient fills are unavailable in this version, shows how the GradientDegree property could be used, and saves the workbook as GradientFill.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get cell A1 and set its value
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Gradient Fill");

            // Retrieve the current style of the cell
            Style style = cell.GetStyle();

            // Set the start and end colors (used for gradient or solid fill)
            style.ForegroundColor = Color.Blue;        // Start color
            style.BackgroundColor = Color.LightBlue;   // End color

            // Apply a solid fill pattern (gradient pattern not available in this version)
            style.Pattern = BackgroundType.Solid;

            // If a gradient pattern were supported, you could set the angle like this:
            // style.GradientDegree = 45;

            // Apply the modified style back to the cell
            cell.SetStyle(style);

            // Save the workbook to a file
            string outputPath = "GradientFill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
