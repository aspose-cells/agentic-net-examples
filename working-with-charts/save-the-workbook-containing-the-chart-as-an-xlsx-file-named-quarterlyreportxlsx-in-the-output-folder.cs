// Title: Generate a column chart in a new Excel workbook and save it as QuarterlyReport.xlsx using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a worksheet, fills month and sales data, adds a column chart, and saves the file as QuarterlyReport.xlsx in an output folder. | Change the chart type to a line chart and rename the output file to SalesTrend.xlsx while preserving the data range and folder logic. | Add a parameter to set the chart title at runtime and allow the caller to specify the destination directory for the saved workbook.
// Common Searches: aspnet c# how to add a column chart to a workbook and export as xlsx with Aspose.Cells | save Aspose.Cells workbook with chart to a custom folder in .NET | example code for creating quarterly sales chart using Aspose.Cells and saving as QuarterlyReport.xlsx | Aspose.Cells C# set chart data range and export workbook to specific directory | create Excel file with chart and specify output path using Aspose.Cells for .NET
// Tags: Aspose.Cells create column chart in workbook | save workbook as XLSX to custom directory | Aspose.Cells set chart data range | C# generate quarterly sales chart with Aspose.Cells | Aspose.Cells output folder handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, adds a worksheet named "Data", populates month and sales values, inserts a column chart titled "Quarterly Sales" using that data, ensures an "output" directory exists, and saves the workbook as "QuarterlyReport.xlsx" in that folder.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr" };
            double[] sales = { 1200, 1500, 1800, 1600 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Quarterly Sales";

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Define the output path (output folder relative to the current directory)
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "output");
            Directory.CreateDirectory(outputFolder); // Ensure the folder exists
            string outputPath = Path.Combine(outputFolder, "QuarterlyReport.xlsx");

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
