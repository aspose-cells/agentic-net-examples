using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class ExportSparklineToPng
{
    static void Main()
    {
        // Define output directory and ensure it exists
        string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
        Directory.CreateDirectory(outputDir);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (A1:D1)
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(15);
        sheet.Cells["D1"].PutValue(30);

        // Define the location cell where the sparkline will be placed (E1)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line and create the sparkline
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];
        Sparkline sparkline = group.Sparklines[0]; // the first (and only) sparkline in the group

        // Configure image options for PNG output
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            Transparent = false
        };

        // Build full file path for the exported image
        string imagePath = Path.Combine(outputDir, "sparkline.png");

        // Export the sparkline to PNG using the ToImage method (string, ImageOrPrintOptions)
        sparkline.ToImage(imagePath, imgOptions);

        // Optionally, save the workbook for reference
        workbook.Save(Path.Combine(outputDir, "WorkbookWithSparkline.xlsx"));

        Console.WriteLine($"Sparkline image saved to: {imagePath}");
    }
}