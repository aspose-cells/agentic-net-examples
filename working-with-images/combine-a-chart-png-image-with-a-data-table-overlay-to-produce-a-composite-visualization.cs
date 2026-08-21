// Title: Create a Composite PNG of a Column Chart and Its Data Table with Aspose.Cells for .NET (C#)
// Description: The sample builds a workbook, populates cells A1:B4 with categories and values, adds a column chart that references the data range, and uses SheetRender with ImageOrPrintOptions to export the whole worksheet—including both the chart and the data table—as a single PNG file. It also shows how to create target folders and handle errors gracefully.
// Keywords: Aspose.Cells | C# | render chart to PNG | SheetRender | ImageOrPrintOptions | export worksheet as image | composite chart image | combine chart and table | Excel to PNG .NET | chart image generation
// Common Searches: export Excel chart with data table to PNG using Aspose.Cells | Aspose.Cells C# render worksheet containing a chart to a single image | combine chart and data range into one PNG file | SheetRender export chart and table as composite image | how to save Excel chart and its source data as PNG in .NET
// Developer Intent: Generate one PNG file that visually merges a column chart with its underlying data table.
// Use Cases: Add a thumbnail to dashboards that shows both the chart and the raw numbers. | Automate creation of chart images for email reports, web pages, or documentation. | Produce printable composite visuals for presentations or technical manuals.
// AI Prompts: Write C# code using Aspose.Cells to render a worksheet that contains a column chart and its data table into a single PNG image, including folder creation and error handling. | Show an example that populates sample data, creates a column chart, and saves the combined view as one image with SheetRender and ImageOrPrintOptions. | Explain how to configure ImageOrPrintOptions for resolution, page layout, and PNG output when exporting a chart with its data table using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsCompositeChart
{
    // The sample builds a workbook, populates cells A1:B4 with categories and values, adds a column chart that references the data range, and uses SheetRender with ImageOrPrintOptions to export the whole worksheet—including both the chart and the data table—as a single PNG file. It also shows how to create target folders and handle errors gracefully.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook and populate sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample rows
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // 2. Add a column chart that uses the data above
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);               // Values
                chart.NSeries.CategoryData = "A2:A4";           // Categories
                chart.Title.Text = "Sample Column Chart";

                // 3. Render the worksheet (including the chart) to a PNG file
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; explicit setting omitted to avoid API mismatch
                    OnePagePerSheet = true
                };
                SheetRender renderer = new SheetRender(sheet, renderOptions);
                string compositePath = "CompositeChart.png";

                // Ensure the directory exists before saving
                string compositeDir = Path.GetDirectoryName(compositePath);
                if (!string.IsNullOrEmpty(compositeDir) && !Directory.Exists(compositeDir))
                {
                    Directory.CreateDirectory(compositeDir);
                }

                try
                {
                    renderer.ToImage(0, compositePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error rendering chart image: {ex.Message}");
                }

                // 4. Save the workbook for reference (optional)
                string workbookPath = "ChartWorkbook.xlsx";

                // Ensure the directory exists before saving
                string workbookDir = Path.GetDirectoryName(workbookPath);
                if (!string.IsNullOrEmpty(workbookDir) && !Directory.Exists(workbookDir))
                {
                    Directory.CreateDirectory(workbookDir);
                }

                try
                {
                    workbook.Save(workbookPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
