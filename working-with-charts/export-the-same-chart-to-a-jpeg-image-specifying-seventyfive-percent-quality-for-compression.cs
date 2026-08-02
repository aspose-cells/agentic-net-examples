// Title: Export Aspose.Cells Column Chart to JPEG with 75% Quality (C#)
// Description: Creates a workbook, adds a column chart, sets ImageOrPrintOptions.Quality to 75, and saves the chart as a JPEG image using chart.ToImage. The workbook can also be saved as XLSX.
// Keywords: Aspose.Cells | C# chart export | JPEG compression | ImageOrPrintOptions Quality | chart.ToImage | column chart image | .NET Excel chart to image | reduce chart file size | web dashboard chart | email report chart
// Common Searches: Aspose.Cells export chart to JPEG with quality setting | C# ImageOrPrintOptions Quality property example | How to save Aspose.Cells chart as JPEG | Set JPEG compression level for chart image in Aspose.Cells | Export multiple charts to JPEG using Aspose.Cells .NET
// Developer Intent: Generate a JPEG file of a column chart with 75% compression quality using Aspose.Cells for .NET.
// Use Cases: Embedding lightweight chart images in web pages to improve load times. | Attaching size‑controlled JPEG charts to automated email reports. | Creating dashboard visuals where a specific image quality balance is required.
// AI Prompts: Show C# code to export an Aspose.Cells chart to PNG with default settings. | Explain how the Quality property of ImageOrPrintOptions affects JPEG size and visual fidelity in Aspose.Cells. | Provide a script that loops through all charts in a worksheet and saves each as a JPEG with different quality percentages.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a column chart, sets ImageOrPrintOptions.Quality to 75, and saves the chart as a JPEG image using chart.ToImage. The workbook can also be saved as XLSX.
class ExportChartToJpeg
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure image options: JPEG format with 75% quality
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Quality range: 0-100
                Quality = 75
                // Image format is inferred from the file extension (.jpg)
            };

            // Export the chart to a JPEG image using the specified options
            chart.ToImage("ChartOutput.jpg", options);

            // Save the workbook (optional)
            workbook.Save("ChartWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
