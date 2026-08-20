// Title: Aspose.Cells for .NET: Create a Pie Chart with a Custom “Other” Slice and Export as 300 DPI PNG
// Description: This .NET example shows how to build a workbook, add category and value data, generate a pie chart, group slices below a configurable percentage into an automatically created “Other” slice, rename that slice with SettableChartGlobalizationSettings, and render the chart to a high‑resolution 300 DPI PNG file. The workbook can also be saved for reference.
// Keywords: Aspose.Cells | C# pie chart | custom other slice | ChartSplitType.PercentValue | high resolution PNG | 300 DPI export | SettableChartGlobalizationSettings | chart rendering Aspose.Cells | export chart image .NET | globalization settings chart
// Common Searches: Aspose.Cells set custom label for Other slice | pie chart split by percent value Aspose | export Aspose.Cells chart to high DPI PNG | change Other slice name in Aspose.Cells | C# render chart as 300 DPI image | globalization settings for charts Aspose.Cells
// Developer Intent: Generate a pie chart that consolidates small categories into a labeled “Other” slice and save the visual as a 300 DPI PNG image.
// Use Cases: Print‑ready sales distribution graphics for marketing brochures. | Automated dashboard thumbnails for web portals. | Localized financial reports where the “Other” label must be translated. | Batch processing of Excel files to create high‑resolution chart images for PDF generation.
// AI Prompts: Modify the sample to use a 10 % split threshold and rename the automatic slice to "Other Categories". | Provide code that exports the same pie chart as a 600 DPI JPEG while preserving the custom "Other" label. | Show how to apply a distinct fill color to the automatically generated "Other" slice after setting its custom name.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // This .NET example shows how to build a workbook, add category and value data, generate a pie chart, group slices below a configurable percentage into an automatically created “Other” slice, rename that slice with SettableChartGlobalizationSettings, and render the chart to a high‑resolution 300 DPI PNG file. The workbook can also be saved for reference.
    public class PieChartCustomOtherLabel
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(15);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the data range to the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Use PercentValue split so that small slices are grouped into an "Other" slice
            chart.NSeries[0].SplitType = ChartSplitType.PercentValue;
            chart.NSeries[0].SplitValue = 20; // values < 20% will be placed in the "Other" slice

            // Set a custom name for the automatically generated "Other" label
            SettableChartGlobalizationSettings globalization = new SettableChartGlobalizationSettings();
            globalization.SetOtherName("Miscellaneous Items");

            // (Optional) Verify the custom name
            string otherName = globalization.GetOtherName();
            Console.WriteLine("Custom 'Other' label set to: " + otherName);

            // Prepare high‑resolution image options (300 DPI PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300
                // ImageFormat defaults to PNG, so no explicit setting required
            };

            // Render the chart to a high‑resolution PNG file
            string pngPath = "PieChart_HighRes.png";
            try
            {
                chart.ToImage(pngPath, imgOptions);
                Console.WriteLine($"Chart image saved to: {Path.GetFullPath(pngPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save chart image: {ex.Message}");
            }

            // Save the workbook (optional, just to keep the file complete)
            string workbookPath = "PieChartWorkbook.xlsx";
            try
            {
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
