// Title: Async Excel Column Chart Generation with Aspose.Cells in WinForms (C#)
// Description: Creates a workbook, fills sample data, adds a column chart, renders the chart to PNG, and saves the file—all executed on a background thread with Task.Run and async I/O to keep the WinForms UI responsive.
// Keywords: Aspose.Cells | async chart creation | C# WinForms | Task.Run | non‑blocking UI | Excel chart image | background thread | .NET | chart generation | workbook creation
// Common Searches: Aspose.Cells generate chart without freezing UI | async Excel chart creation C# WinForms | Task.Run Aspose.Cells chart example | save Excel chart as PNG asynchronously | non‑blocking workbook generation Aspose.Cells
// Developer Intent: Build an Excel workbook with a column chart and export the chart image using asynchronous code so the UI remains responsive.
// Use Cases: Generate large data‑driven charts on a background thread while showing a progress indicator in a WinForms form. | Create a preview PNG of a chart during workbook construction, then deliver both the image and the .xlsx file to the user. | Expose an API that assembles a chart‑filled workbook asynchronously, freeing the request thread for other operations.
// AI Prompts: Write an async method that builds a pie chart with Aspose.Cells, saves it as PNG, and returns the workbook bytes. | Show how to report progress from a Task.Run chart creation to a WinForms ProgressBar using IProgress. | Provide robust error handling for asynchronous chart generation with Aspose.Cells, including logging and user‑friendly messages.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAsyncChartDemo
{
    // Creates a workbook, fills sample data, adds a column chart, renders the chart to PNG, and saves the file—all executed on a background thread with Task.Run and async I/O to keep the WinForms UI responsive.
    class Program
    {
        // Entry point with async support
        static async Task Main(string[] args)
        {
            try
            {
                Console.WriteLine("Generating chart asynchronously...");

                // Run chart creation on a background thread
                var workbookBytes = await Task.Run(() => CreateChart());

                // Save the workbook
                string workbookPath = "AsyncChartWorkbook.xlsx";
                await File.WriteAllBytesAsync(workbookPath, workbookBytes);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a workbook with a chart, saves the chart image, and returns the workbook bytes
        private static byte[] CreateChart()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                var chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Configure chart appearance
                chart.Title.Text = "Async Generated Chart";
                chart.ShowLegend = true;
                chart.SizeWithWindow = true;

                // Ensure layout is calculated before rendering
                chart.Calculate();

                // Save chart image to file
                string imagePath = "AsyncChart.png";
                chart.ToImage(imagePath);
                Console.WriteLine($"Chart image saved to: {Path.GetFullPath(imagePath)}");

                // Save workbook to a memory stream and return bytes
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                // Propagate exception to caller
                throw new InvalidOperationException("Failed to create chart and workbook.", ex);
            }
        }
    }
}
