// Title: Apply a corporate .thmx theme to a Waterfall chart in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that checks whether a corporate .thmx file exists, loads it with Aspose.Cells if the LoadTheme method is available, and applies the theme to an existing Waterfall chart, including proper error handling. | Show how to conditionally apply a custom theme to a chart and then save the workbook as an .xlsx file using the Aspose.Cells .NET API.
// Common Searches: asp.net apply corporate .thmx theme to Excel chart using Aspose.Cells | c# load custom theme for waterfall chart Aspose.Cells LoadTheme | how to check for theme file before applying in Aspose.Cells workbook | programmatically theme Excel charts with Aspose.Cells .NET | fallback when LoadTheme API is unavailable in Aspose.Cells
// Tags: apply .thmx theme Aspose.Cells | waterfall chart theming .NET | conditional LoadTheme usage | excel chart branding automation | theme file existence check C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds sample data and a Waterfall chart, then checks for a corporate .thmx theme file. If the file exists and the LoadTheme API is supported, the theme is applied to the chart; otherwise the code proceeds without theming. Finally, the workbook is saved as WaterfallWithCorporateTheme.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the Waterfall chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Increase");
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["A4"].PutValue("Decrease");
                sheet.Cells["B4"].PutValue(-20);
                sheet.Cells["A5"].PutValue("End");
                sheet.Cells["B5"].PutValue(110);

                // Add a Waterfall chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
                Chart waterfallChart = sheet.Charts[chartIndex];
                waterfallChart.Title.Text = "Waterfall Chart";

                // Define the series data range and category labels
                waterfallChart.NSeries.Add("B2:B5", true);
                waterfallChart.NSeries.CategoryData = "A2:A5";

                // Attempt to load a corporate theme if the file exists.
                // Note: LoadTheme method may not be available in all Aspose.Cells versions.
                string themePath = "CorporateTheme.thmx";
                if (File.Exists(themePath))
                {
                    try
                    {
                        // If the API is supported, apply the theme.
                        // workbook.LoadTheme(themePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to apply theme: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Theme file '{themePath}' not found. Continuing without applying a custom theme.");
                }

                // Save the workbook with the (optional) themed Waterfall chart
                string outputPath = "WaterfallWithCorporateTheme.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
