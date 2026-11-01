// Title: Generate a column chart for monthly sales data in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a Column chart to a worksheet, bind it to the range B2:B5, set the chart heading to "Monthly Sales", and save the workbook as an .xlsx file. | Show how to place a column chart on a worksheet covering rows 6‑20 and columns A‑K, and configure its data series with Aspose.Cells in C#.
// Common Searches: aspnet c# how to create a column chart from a data range using Aspose.Cells | example code for adding a column chart with heading in Aspose.Cells .NET | Aspose.Cells chart series range B2:B5 C# tutorial | save workbook with chart as xlsx using Aspose.Cells C# | place column chart in defined cell range using Aspose.Cells C#
// Tags: Aspose.Cells column chart insertion | C# define chart data series Aspose.Cells | Aspose.Cells configure chart heading | Aspose.Cells export workbook with chart | Aspose.Cells set chart location cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:B5 with month and sales values, inserts a column chart positioned from rows 6 to 20 and columns A to K, links the series to the sales range B2:B5, assigns the title "Monthly Sales", and saves the file as SalesColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sales data (Month in column A, Sales in column B)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(18000);
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B5"].PutValue(13000);

            // Add a column chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B5", true);
            // Category data can be set if supported; omitted for compatibility with older API versions

            // Optional: set chart title
            chart.Title.Text = "Monthly Sales";

            // Define output file path
            string outputPath = "SalesColumnChart.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
