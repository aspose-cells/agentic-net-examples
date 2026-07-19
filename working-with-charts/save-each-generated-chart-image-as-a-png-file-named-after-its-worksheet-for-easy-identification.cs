// Title: Save Worksheet Charts as PNG Files Named After Their Sheets with Aspose.Cells (C#)
// Description: Creates a workbook with two sheets, adds a column chart to each, then loops through the worksheets and exports the first chart of each sheet to a PNG file whose name matches the worksheet (e.g., Sheet1.png). The workbook is then saved.
// Keywords: Aspose.Cells | C# chart export | export chart to PNG | save chart image by worksheet name | Aspose.Cells ToImage | multiple worksheet chart export | chart image naming | Aspose.Cells .NET | Excel chart PNG | automate chart export
// Common Searches: How to export chart from each worksheet as PNG using Aspose.Cells C# | Aspose.Cells save chart image with sheet name | Export multiple charts to separate PNG files Aspose.Cells | C# code to loop through worksheets and save charts as images | Aspose.Cells ToImage example for column chart
// Developer Intent: Export every chart in a workbook to a PNG file whose filename matches its worksheet.
// Use Cases: Create chart images for reporting dashboards | Generate assets for web pages or documentation | Batch export chart visuals for presentation decks | Automate image creation for data‑driven newsletters | Provide chart snapshots for API responses
// AI Prompts: Write C# code that uses Aspose.Cells to export all charts in a workbook to JPEG files named with the worksheet and chart index. | Explain how to set image resolution, DPI, and format when exporting charts with Aspose.Cells. | Show how to handle worksheets containing multiple charts and produce unique filenames like Sheet1_Chart1.png. | Demonstrate saving chart images to a memory stream instead of disk using Aspose.Cells. | Provide a PowerShell script that calls a .NET assembly to export worksheet charts as PNG files.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExport
{
    // Creates a workbook with two sheets, adds a column chart to each, then loops through the worksheets and exports the first chart of each sheet to a PNG file whose name matches the worksheet (e.g., Sheet1.png). The workbook is then saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains one default worksheet)
                Workbook workbook = new Workbook();

                // Add two worksheets with sample data and a chart in each
                for (int i = 0; i < 2; i++)
                {
                    Worksheet sheet;

                    // The first worksheet already exists; add a new one for the second iteration
                    if (i == 0)
                    {
                        sheet = workbook.Worksheets[0];
                    }
                    else
                    {
                        workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[i];
                    }

                    sheet.Name = $"Sheet{i + 1}";

                    // Populate sample data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["A4"].PutValue("C");

                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["B2"].PutValue(10 + i * 5);
                    sheet.Cells["B3"].PutValue(20 + i * 5);
                    sheet.Cells["B4"].PutValue(30 + i * 5);

                    // Add a column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.SetChartDataRange("A1:B4", true);
                }

                // Export each chart as a PNG file named after its worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Charts.Count > 0)
                    {
                        Chart chart = sheet.Charts[0];
                        string imageFileName = $"{sheet.Name}.png";
                        chart.ToImage(imageFileName);
                    }
                }

                // Save the workbook (optional)
                workbook.Save("ChartsWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
