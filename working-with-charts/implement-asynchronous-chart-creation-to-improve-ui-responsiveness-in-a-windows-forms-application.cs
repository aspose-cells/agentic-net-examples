// Title: Asynchronously generate an Aspose.Cells chart and export it as a PNG image in a WinForms application
// Description: Demonstrates how to create a workbook, add a column chart, and export the chart as a PNG file using Aspose.Cells while running the operation on a background thread to keep the Windows Forms UI responsive.
// Keywords: Aspose.Cells async chart | C# asynchronous chart image | WinForms non‑blocking chart generation | export Excel chart to PNG | background thread Aspose.Cells | Task.Run chart creation | UI responsive Excel chart | C# async file I/O Aspose.Cells
// Common Searches: async Aspose.Cells chart WinForms | how to create Excel chart without freezing UI | Aspose.Cells generate chart image in background | C# export chart to PNG asynchronously | Windows Forms chart generation Aspose.Cells example
// Developer Intent: Create an Excel chart and its PNG image on a background thread so the WinForms UI remains responsive.
// Use Cases: Run chart creation and image export inside Task.Run, then await the task from a button click. | Update a PictureBox with the generated PNG after the async operation completes, using Invoke or SynchronizationContext. | Show a progress bar while the chart is being rendered and saved to the user's desktop in a non‑blocking way. | Integrate the async chart routine into a larger data‑processing pipeline that must not block the UI thread.
// AI Prompts: Write an async version of SaveChartImage that returns Task and can be awaited from a WinForms button click. | Show code that calls the async chart method inside Task.Run and marshals the resulting image path back to the UI thread to display in a PictureBox. | Provide a sample that uses IProgress<T> to report chart‑generation progress while exporting the chart image with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAsyncChartDemo
{
    // Demonstrates how to create a workbook, add a column chart, and export the chart as a PNG file using Aspose.Cells while running the operation on a background thread to keep the Windows Forms UI responsive.
    class Program
    {
        static void Main()
        {
            try
            {
                // Generate chart and save its image to the desktop
                SaveChartImage();

                Console.WriteLine("Chart image and workbook have been saved to the desktop.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void SaveChartImage()
        {
            // Create a new workbook (lifecycle rule: create)
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

            // Add a column chart to the worksheet (rule: ChartCollection.Add)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Optional chart customizations
            chart.Title.Text = "Sample Column Chart";
            chart.ShowLegend = true;
            chart.SizeWithWindow = true;

            // Ensure the chart layout is calculated before rendering
            chart.Calculate();

            // Determine desktop paths
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string imagePath = Path.Combine(desktopPath, "AsyncChartDemo.png");
            string workbookPath = Path.Combine(desktopPath, "AsyncChartDemo.xlsx");

            // Save the chart as an image file
            chart.ToImage(imagePath);

            // Save the workbook to a file (lifecycle rule: save)
            workbook.Save(workbookPath, SaveFormat.Xlsx);
        }
    }
}
