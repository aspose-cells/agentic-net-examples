// Title: Export a Sparkline as a PNG image with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a line sparkline from cells A1:D1, sets 300 DPI PNG options, and uses Sparkline.ToImage to write the sparkline to a specified output folder. The workbook can also be saved alongside the image.
// Keywords: Aspose.Cells | C# | Sparkline export | PNG image | ImageOrPrintOptions | high resolution | ToImage method | line sparkline | .NET Excel | programmatic sparkline image
// Common Searches: Aspose.Cells export sparkline to PNG C# | Save single sparkline as image .NET | Sparkline.ToImage example | High DPI PNG from Excel sparkline | C# code to export sparkline image
// Developer Intent: Export a selected sparkline from an Excel worksheet to a PNG file programmatically.
// Use Cases: Generate thumbnail graphics of sparklines for web dashboards. | Create high‑quality images for PDF or HTML reports without embedding the full workbook. | Archive visual snapshots of key metrics for documentation or compliance.
// AI Prompts: Write C# code that extracts a specific sparkline from an Aspose.Cells worksheet and saves it as a 300 DPI PNG file. | Show how to iterate over all sparkline groups in a workbook and export each sparkline to separate PNG files using Aspose.Cells. | Explain how to adjust PNG export settings such as resolution, transparency, and background color when saving a sparkline with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a line sparkline from cells A1:D1, sets 300 DPI PNG options, and uses Sparkline.ToImage to write the sparkline to a specified output folder. The workbook can also be saved alongside the image.
class ExportSparkline
{
    static void Main()
    {
        // Define the output directory and ensure it exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(15);
        sheet.Cells["D1"].PutValue(30);

        // Define the cell where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a line sparkline group using the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            sheet.Name + "!A1:D1",
            false,
            location);

        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Retrieve the first sparkline from the group
        Sparkline sparkline = group.Sparklines[0];

        // Configure image options for PNG output with high resolution
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            Transparent = false
        };

        // Build the full path for the exported image
        string imagePath = Path.Combine(outputDir, "sparkline.png");

        // Export the sparkline to a PNG file
        sparkline.ToImage(imagePath, imgOptions);

        // (Optional) Save the workbook that contains the sparkline
        workbook.Save(Path.Combine(outputDir, "WorkbookWithSparkline.xlsx"));

        Console.WriteLine($"Sparkline image successfully saved to: {imagePath}");
    }
}
