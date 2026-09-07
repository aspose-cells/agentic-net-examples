// Title: Add a column chart to an Excel worksheet and bind its data source with SetChartDataRange using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook, fills cells A1:B5 with sample data, adds a Column chart, and binds the chart to that range (including headers) using the SetChartDataRange method in Aspose.Cells. | Show how to calculate the last populated row in a worksheet and use that value to set a dynamic data range for a column chart with SetChartDataRange in Aspose.Cells. | Provide an example that assigns a title to the chart, positions it on the sheet, and saves the workbook as an .xlsx file after configuring the column chart with Aspose.Cells.
// Common Searches: aspnet aspose.cells setchartdatarange column chart example c# | how to bind excel column chart to data range with headers using aspose.cells | c# create column chart in workbook and set data source programmatically | aspose.cells set chart data range dynamically based on last row | saving workbook after adding chart with aspose.cells .net
// Tags: create column chart Aspose.Cells C# | SetChartDataRange usage in Aspose.Cells | assign chart data source Excel .NET | chart title assignment Aspose.Cells | save workbook with chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // This program creates a new workbook, populates cells A1:B5 with sample data, adds a column chart at a specified location, binds the chart to the range (including headers) using SetChartDataRange, sets the chart title to "Monthly Sales", and saves the file as ColumnChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
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

            // Add a new column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B5", true);

            // Set a title for the chart
            chart.Title.Text = "Monthly Sales";

            // Define output file path
            string outputPath = "ColumnChart.xlsx";

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
