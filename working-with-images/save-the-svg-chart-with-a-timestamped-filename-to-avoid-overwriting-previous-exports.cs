// Title: Export an Aspose.Cells column chart to a PNG file with a timestamped filename in C#
// AI Prompts: Generate C# code that creates a workbook, adds a column chart, and saves the chart as a PNG file whose name includes the current date and time. | Write C# logic to verify that the target folder exists and create it if necessary before exporting the chart image with Aspose.Cells. | Adapt the example to export the chart as an SVG file while preserving the timestamped naming convention.
// Common Searches: Aspose.Cells C# export chart to PNG with unique timestamp filename | how to create timestamped file name for chart image using Aspose.Cells | C# save Aspose.Cells chart as SVG without overwriting previous files | ensure output directory exists before saving chart image Aspose.Cells | generate column chart and export as image with date-time suffix in C#
// Tags: export chart to PNG with Aspose.Cells | timestamped filename for Aspose.Cells image export | ensure output directory exists Aspose.Cells C# | save chart as SVG using Aspose.Cells ImageOrPrintOptions | column chart image generation Aspose.Cells

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The program creates a workbook, fills it with sample data, adds a column chart, and exports the chart as a PNG image whose filename includes a yyyyMMdd_HHmmss timestamp. It also ensures the destination directory exists before writing the file.
class Program
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories
            chart.Title.Text = "Sample Chart";

            // Recalculate formulas to ensure chart data is up‑to‑date
            workbook.CalculateFormula();

            // Build a timestamped filename to avoid overwriting previous exports
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"Chart_{timestamp}.png";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(fileName));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the chart as an image (PNG)
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; no need to set ImageFormat property
                    OnePagePerSheet = true
                };
                chart.ToImage(fs, imgOptions);
            }

            Console.WriteLine($"Chart saved successfully as {fileName}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
