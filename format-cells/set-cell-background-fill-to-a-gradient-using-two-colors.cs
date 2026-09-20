// Title: Apply a two‑color background fill to a cell in Excel using Aspose.Cells for .NET (solid pattern fallback)
// AI Prompts: Write C# code with Aspose.Cells that sets a blue foreground and light‑blue background on cell A1 and saves the workbook as GradientFill.xlsx. | Show how to simulate a gradient fill in Aspose.Cells by configuring foreground and background colors and using a solid pattern as a fallback. | Create a reusable C# method that accepts two System.Drawing.Color values and applies them as a gradient‑style fill to a specified worksheet cell using Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to set cell background with two colors | C# Aspose.Cells simulate gradient fill for Excel cell | apply foreground and background colors to a cell using Aspose.Cells style | workaround for missing gradient fill support in Aspose.Cells | save Excel file with custom cell background colors using Aspose.Cells C#
// Tags: Aspose.Cells set cell background colors C# | gradient fill fallback Aspose.Cells | cell style foreground background Aspose.Cells | Excel workbook save Aspose.Cells C# | solid pattern as gradient workaround Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program creates a new workbook, writes "Gradient Fill" to cell A1, defines a style with a blue foreground and light‑blue background, sets the pattern to Solid (since gradient fill isn’t directly supported), applies the style to the cell, and saves the file as GradientFill.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Access cell A1 and set a value
            var cell = worksheet.Cells["A1"];
            cell.PutValue("Gradient Fill");

            // Create a new style
            var style = workbook.CreateStyle();

            // Set gradient colors (foreground and background)
            style.ForegroundColor = Color.Blue;
            style.BackgroundColor = Color.LightBlue;

            // Apply a solid pattern (gradient fill is not directly supported in this API version)
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the cell
            cell.SetStyle(style);

            // Define output file path
            string outputPath = "GradientFill.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
