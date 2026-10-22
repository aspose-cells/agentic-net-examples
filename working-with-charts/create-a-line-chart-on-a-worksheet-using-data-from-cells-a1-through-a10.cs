// Title: Create a line chart from cells A1‑A10 on a worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a line chart to the first worksheet, binds its series to the range A1:A10, sets a chart title, and saves the workbook using Aspose.Cells. | Show how to populate column A with sample values, position a line chart (rows 5‑25, columns 0‑10), and define its data source in Aspose.Cells. | Provide a step‑by‑step example of creating and customizing a line chart in an Excel file with Aspose.Cells, including data population, series definition, title assignment, and file saving.
// Common Searches: Aspose.Cells C# example for creating a line chart from column A data | How to bind a line chart series to A1:A10 using Aspose.Cells .NET | C# code to add and position a line chart in an Excel worksheet with Aspose.Cells | Set chart title and save workbook with chart using Aspose.Cells for .NET | Create line chart in Excel file programmatically with Aspose.Cells C# tutorial
// Tags: Aspose.Cells create line chart | Aspose.Cells set chart series range | Aspose.Cells add chart to worksheet | Aspose.Cells line chart positioning | Aspose.Cells save workbook with chart | C# Aspose.Cells line chart example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a new workbook, fills cells A1‑A10 with values 1‑10, adds a line chart positioned from row 5 column 0 to row 25 column 10, binds the chart series to the range A1:A10, assigns a title, and saves the file as LineChart.xlsx.
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

            // Populate cells A1 through A10 with sample data
            for (int i = 0; i < 10; i++)
            {
                // Row index i, column index 0 corresponds to column A
                sheet.Cells[i, 0].PutValue(i + 1); // A1=1, A2=2, ..., A10=10
            }

            // Add a line chart to the worksheet (positioned from row 5, column 0 to row 25, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart series (A1:A10)
            chart.NSeries.Add("A1:A10", true);

            // Optional: set a title for the chart
            chart.Title.Text = "Line Chart Example";

            // Save the workbook to a file
            workbook.Save("LineChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
