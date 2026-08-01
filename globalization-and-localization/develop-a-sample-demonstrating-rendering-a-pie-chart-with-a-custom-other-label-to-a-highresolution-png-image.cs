// Title: C# – Render a High‑Resolution PNG Pie Chart with a Custom “Other” Label Using Aspose.Cells
// Description: This example creates a workbook, fills it with category/value data, adds a pie chart, applies a custom name to the automatically generated “Other” slice via SettableChartGlobalizationSettings, sets ImageOrPrintOptions to 300 DPI, and exports the chart as a high‑resolution PNG. The workbook can also be saved for later editing.
// Keywords: Aspose.Cells | C# | .NET | pie chart | high resolution PNG | 300 DPI | custom Other label | SettableChartGlobalizationSettings | chart globalization | export chart image | localization
// Common Searches: Aspose.Cells custom Other label pie chart | export pie chart as high DPI PNG C# | SettableChartGlobalizationSettings example | render Aspose.Cells chart to high resolution image | globalize chart labels Aspose.Cells
// Developer Intent: Generate a pie chart, assign a localized "Other" slice name, and export the chart as a 300 DPI PNG image with Aspose.Cells for .NET.
// Use Cases: Create printable reports that require a crisp PNG pie chart with minor categories grouped under a translated "Other" label. | Supply high‑resolution chart images for dashboards, presentations, or documentation while preserving the original workbook. | Automate multilingual chart generation where the "Other" slice text must be customized per locale.
// AI Prompts: Show how to set the DPI for a chart image exported with Aspose.Cells in C#. | Provide code to change the "Other" slice label of a pie chart using SettableChartGlobalizationSettings. | Explain how to batch‑export multiple charts from a workbook to high‑resolution PNG files with Aspose.Cells.

using System;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System.IO;

namespace AsposeCellsPieChartHighRes
{
    // This example creates a workbook, fills it with category/value data, adds a pie chart, applies a custom name to the automatically generated “Other” slice via SettableChartGlobalizationSettings, sets ImageOrPrintOptions to 300 DPI, and exports the chart as a high‑resolution PNG. The workbook can also be saved for later editing.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["A5"].PutValue("Grape");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(70);

                // Add a pie chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 12);
                Chart pieChart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                pieChart.NSeries.Add("B2:B5", true);
                pieChart.NSeries.CategoryData = "A2:A5";

                // Set a custom name for the "Other" label using globalization settings
                SettableChartGlobalizationSettings globalizationSettings = new SettableChartGlobalizationSettings();
                globalizationSettings.SetOtherName("Miscellaneous Items");
                Console.WriteLine("Custom Other label: " + globalizationSettings.GetOtherName());

                // Configure high‑resolution image options (e.g., 300 DPI)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                    // ImageFormat defaults to PNG when the file extension is .png
                };

                // Ensure the output directory exists
                string outputImagePath = "HighResPieChart.png";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputImagePath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Render the chart to a high‑resolution PNG file
                pieChart.ToImage(outputImagePath, imgOptions);
                Console.WriteLine($"Pie chart rendered to high‑resolution PNG: {outputImagePath}");

                // Save the workbook (optional, to keep the chart in the file)
                string workbookPath = "PieChartWorkbook.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
