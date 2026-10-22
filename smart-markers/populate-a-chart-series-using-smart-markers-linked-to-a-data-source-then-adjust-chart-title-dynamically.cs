// Title: Generate a column chart from a DataTable using smart markers and update its title dynamically with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a DataTable into a worksheet, adds a column chart using smart markers, and sets the chart title to include the row count. | Demonstrate how to bind a chart series to a worksheet range with smart markers and change the chart title at runtime in Aspose.Cells. | Provide a complete example that saves the workbook as an .xlsx file after configuring the smart‑marker chart and dynamic title.
// Common Searches: aspnet how to bind column chart series to a DataTable with smart markers in Aspose.Cells | c# set Excel chart title based on number of rows using Aspose.Cells | using smart markers to populate chart data range Aspose.Cells .NET example | dynamic chart title after populating worksheet data Aspose.Cells C# | save workbook as xlsx after creating smart marker chart Aspose.Cells
// Tags: smart markers chart series Aspose.Cells C# | dynamic chart title Aspose.Cells | bind column chart to DataTable Aspose.Cells | save workbook as xlsx Aspose.Cells | populate worksheet from DataTable Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, writes a DataTable of months and sales to a worksheet, adds a column chart whose series are defined with smart markers, updates the chart title to show the total row count, and saves the file as SmartMarkerChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Prepare a data source (DataTable) with sample data
            DataTable dt = new DataTable();
            dt.Columns.Add("Month", typeof(string));
            dt.Columns.Add("Sales", typeof(double));
            dt.Rows.Add("Jan", 1200);
            dt.Rows.Add("Feb", 1500);
            dt.Rows.Add("Mar", 1800);
            dt.Rows.Add("Apr", 1300);
            dt.Rows.Add("May", 1700);

            // Write the data into the worksheet starting at cell A1
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Sales");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.Cells[i + 1, 0].PutValue(dt.Rows[i]["Month"]);
                ws.Cells[i + 1, 1].PutValue(dt.Rows[i]["Sales"]);
            }

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 7, 0, 27, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.Title.Text = "Placeholder Title";

            // Configure the chart series using smart markers (direct range references)
            chart.NSeries.Add("=Data!B2:B6", true);
            chart.NSeries[0].Name = "=Data!B1";
            chart.NSeries[0].XValues = "=Data!A2:A6";

            // Dynamically adjust the chart title after data has been populated
            chart.Title.Text = $"Sales Report - Total Items: {dt.Rows.Count}";

            // Save the workbook to a file
            string outputPath = "SmartMarkerChart.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
