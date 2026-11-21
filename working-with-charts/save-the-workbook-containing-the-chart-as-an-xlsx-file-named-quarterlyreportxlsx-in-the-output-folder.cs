// Title: Generate a column chart with Aspose.Cells and save the workbook as QuarterlyReport.xlsx in an output directory (C#)
// Common Searches: Aspose.Cells C# create column chart and save workbook to specific folder | how to export Aspose.Cells workbook containing a chart to QuarterlyReport.xlsx | C# code to save Aspose.Cells workbook with chart in output folder | Aspose.Cells save workbook as XLSX with chart in subdirectory | create chart in Aspose.Cells and write file to output directory using C#
// Tags: Aspose.Cells column chart creation C# | Aspose.Cells save workbook as XLSX with chart | Aspose.Cells output folder handling C# | Aspose.Cells chart data series assignment | Aspose.Cells workbook.Save to custom path | Aspose.Cells generate quarterly sales report

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a worksheet named "Data", populates month and sales data, inserts a column chart titled "Quarterly Sales", ensures an "output" directory exists, and saves the workbook as "QuarterlyReport.xlsx" in that folder using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            string[] months = { "Jan", "Feb", "Mar", "Apr" };
            double[] sales = { 12000, 15000, 13000, 17000 };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Quarterly Sales";

            // Set the data range for the chart (values)
            chart.NSeries.Add("B2:B5", true);
            // Category data (months) is automatically taken from the first column when using the overload above,
            // so explicit CategoryData assignment is not required.

            // Define the output path (output folder relative to the current directory)
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "output");
            Directory.CreateDirectory(outputFolder); // Ensure the folder exists
            string outputPath = Path.Combine(outputFolder, "QuarterlyReport.xlsx");

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
