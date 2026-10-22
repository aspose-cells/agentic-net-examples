// Title: Asynchronously create a column chart in an Excel workbook for a WinForms application using Aspose.Cells for .NET
// AI Prompts: Generate an async method that builds a Workbook, inserts sample data, adds a Column chart, and saves the file as .xlsx with Aspose.Cells. | Modify the async chart routine to accept a data range parameter and produce a Line chart instead of a Column chart. | Integrate IProgress reporting into the async chart creation to update a WinForms progress bar while the workbook is being generated.
// Common Searches: how to generate Excel charts without blocking UI using Aspose.Cells C# | async column chart example for WinForms with Aspose.Cells | Aspose.Cells create chart on background thread and save workbook | improve responsiveness when adding charts to Excel in .NET application | using Task.Run to build Excel workbook with chart Aspose.Cells
// Tags: asynchronous chart generation Aspose.Cells .NET | column chart creation with Aspose.Cells C# | save workbook async Aspose.Cells | WinForms UI responsiveness Excel chart | Task.Run Aspose.Cells chart example | background thread Excel chart generation

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Provides an async method that creates a Workbook, populates sample data, adds a column chart, ensures the output folder exists, and saves the workbook as an .xlsx file using Aspose.Cells, enabling UI‑responsive chart generation in WinForms.
    public class ChartHelper
    {
        // Asynchronously creates a workbook, adds data and a chart, then saves the file.
        public async Task CreateChartAsync(string outputFilePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    // 1. Create a new workbook.
                    Workbook workbook = new Workbook();

                    // 2. Access the first worksheet.
                    Worksheet sheet = workbook.Worksheets[0];

                    // 3. Populate sample data.
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["A2"].PutValue("Jan");
                    sheet.Cells["A3"].PutValue("Feb");
                    sheet.Cells["A4"].PutValue("Mar");
                    sheet.Cells["A5"].PutValue("Apr");
                    sheet.Cells["B2"].PutValue(120);
                    sheet.Cells["B3"].PutValue(150);
                    sheet.Cells["B4"].PutValue(180);
                    sheet.Cells["B5"].PutValue(200);

                    // 4. Add a column chart.
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.Title.Text = "Monthly Sales";

                    // 5. Add a series (values) with categories.
                    int seriesIndex = chart.NSeries.Add("B2:B5", true);
                    chart.NSeries[seriesIndex].Name = "Sales";

                    // 6. Ensure the output directory exists.
                    string dir = Path.GetDirectoryName(outputFilePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    // 7. Save the workbook.
                    workbook.Save(outputFilePath, SaveFormat.Xlsx);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error creating chart: {ex.Message}");
                    throw;
                }
            });
        }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            string outputPath = "ChartReport.xlsx";

            try
            {
                await new ChartHelper().CreateChartAsync(outputPath);
                Console.WriteLine($"Chart created and saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to create chart: {ex.Message}");
            }
        }
    }
}
