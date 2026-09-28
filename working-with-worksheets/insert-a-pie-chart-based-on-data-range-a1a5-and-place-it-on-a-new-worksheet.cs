// Title: Create a new worksheet and insert a pie chart from cells A1:A5 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a new worksheet, places a pie chart on it, links the chart series to the range A1:A5 on the first sheet, sets a title, and saves the workbook. | Demonstrate how to programmatically bind a pie chart series to a cell range and position the chart on a separate sheet in an Aspose.Cells workbook.
// Common Searches: aspnet cells how to add a pie chart on a separate worksheet in C# | c# Aspose.Cells create pie chart from range A1:A5 on new sheet | insert chart into new worksheet using Aspose.Cells library
// Tags: Aspose.Cells create pie chart on new worksheet | Aspose.Cells bind chart series to cell range | Aspose.Cells set chart title programmatically | Aspose.Cells save workbook with chart | C# add worksheet and chart using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills cells A1‑A5 with sample data, adds a new worksheet named "PieChart", inserts a pie chart at position (0,0) sized 400×300, binds the chart series to Sheet1!A1:A5, sets a chart title, and saves the file as PieChartWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate sample data in the first worksheet (A1:A5)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A1"].PutValue(10);
            dataSheet.Cells["A2"].PutValue(20);
            dataSheet.Cells["A3"].PutValue(30);
            dataSheet.Cells["A4"].PutValue(25);
            dataSheet.Cells["A5"].PutValue(15);

            // Add a new worksheet to host the pie chart
            int chartSheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "PieChart";

            // Insert a pie chart on the new worksheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Pie, 0, 0, 400, 300);
            Chart pieChart = chartSheet.Charts[chartIndex];

            // Define the data range for the chart series (A1:A5 from the first sheet)
            pieChart.NSeries.Add("Sheet1!A1:A5", true);

            // Optional: set a title for the chart
            pieChart.Title.Text = "Sample Pie Chart";

            // Save the workbook to a file
            string outputPath = "PieChartWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
