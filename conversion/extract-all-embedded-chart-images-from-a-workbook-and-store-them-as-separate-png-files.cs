// Title: Extract Embedded Excel Chart Images to PNG with Aspose.Cells for .NET
// Description: Loads an .xlsx file, creates an output folder, loops through each worksheet and its charts, and saves every chart as a uniquely‑named PNG using Chart.ToImage, while reporting the total count.
// Keywords: Aspose.Cells | C# | extract chart images | Excel chart to PNG | Chart.ToImage | export all charts | save chart as image | workbook chart extraction | Aspose.Cells example | GitHub
// Common Searches: export all charts from Excel to PNG C# | Aspose.Cells save chart as image example | how to extract chart images from .xlsx using .NET | C# code to loop worksheets and export charts | Aspose.Cells Chart.ToImage usage
// Developer Intent: Programmatically retrieve every chart in an Excel workbook and write each one to a separate PNG file.
// Use Cases: Create image assets for reports or presentations from workbook charts. | Build a thumbnail gallery of all charts for a dashboard or documentation site. | Archive visual representations of charts before performing bulk workbook modifications.
// AI Prompts: Generate C# code with Aspose.Cells that extracts all workbook charts and saves them as PNG files using a custom naming pattern. | Show how to modify the sample to export charts in JPEG format and prepend the chart title to the filename. | Explain how to add error handling for workbooks that contain no charts or unsupported chart types.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an .xlsx file, creates an output folder, loops through each worksheet and its charts, and saves every chart as a uniquely‑named PNG using Chart.ToImage, while reporting the total count.
class ExtractChartImages
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Directory where extracted chart images will be saved
        string outputFolder = "ChartImages";
        Directory.CreateDirectory(outputFolder);

        // Load the workbook (uses the provided load rule)
        Workbook workbook = new Workbook(workbookPath);

        int extractedCount = 0;

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the current worksheet
            for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
            {
                Chart chart = sheet.Charts[chartIndex];

                // Build a unique file name for each chart image
                string imageFileName = $"Chart_{sheet.Name}_{chartIndex}.png";
                string imagePath = Path.Combine(outputFolder, imageFileName);

                // Save the chart as a PNG image (uses the provided ToImage rule)
                chart.ToImage(imagePath, ImageType.Png);

                extractedCount++;
            }
        }

        Console.WriteLine($"Total chart images extracted: {extractedCount}");
    }
}
