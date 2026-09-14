// Title: Export a specific line sparkline to a high‑resolution PNG file with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a line sparkline from a range and saves the sparkline as a 300 dpi PNG using Aspose.Cells. | Show how to use Aspose.Cells' Sparkline.ToImage method to render a single sparkline to a PNG file in a custom output folder. | Generate a method that accepts a Sparkline object and ImageOrPrintOptions, then exports the sparkline to a PNG image with the specified resolution.
// Common Searches: Aspose.Cells export individual sparkline as PNG in C# | How to render a single sparkline to an image file using Aspose.Cells for .NET | C# code to save a line sparkline to high resolution PNG with Aspose.Cells | Export sparkline from workbook to PNG folder Aspose.Cells example
// Tags: sparkline toimage png export Aspose.Cells | line sparkline image rendering .NET | custom output directory sparkline PNG Aspose | imageorprintoptions high dpi sparkline export | aspnet sparkline png generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a line sparkline linked to cells A1:D1, configures ImageOrPrintOptions for 300 dpi PNG, and uses Sparkline.ToImage to save the sparkline as a PNG file in an output folder, then saves the workbook.
class ExportSparkline
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data that the sparkline will represent
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(15);
        sheet.Cells["D1"].PutValue(30);

        // Define the cell where the sparkline will be placed (E1)
        CellArea sparkLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a line sparkline group linked to the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            sheet.Name + "!A1:D1",
            false,
            sparkLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Retrieve the first sparkline from the group
        Sparkline sparkline = group.Sparklines[0];

        // Configure image options for PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };

        // Ensure the output directory exists
        string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
        Directory.CreateDirectory(outputDir);

        // Export the sparkline as a PNG image
        string sparkImagePath = Path.Combine(outputDir, "sparkline.png");
        sparkline.ToImage(sparkImagePath, imgOptions);

        // Optionally save the workbook for reference
        string workbookPath = Path.Combine(outputDir, "workbook.xlsx");
        workbook.Save(workbookPath);

        Console.WriteLine($"Sparkline image saved to: {sparkImagePath}");
        Console.WriteLine($"Workbook saved to: {workbookPath}");
    }
}
