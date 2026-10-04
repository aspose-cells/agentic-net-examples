// Title: Generate a radar chart in C# with Aspose.Cells and display category names on the radial axis
// AI Prompts: Write C# code using Aspose.Cells to create a radar chart, bind categories from column A and values from column B, and show the category names on the radial axis. | Demonstrate how to add a series with a combined range and assign a chart title for a radar chart in Aspose.Cells for .NET.
// Common Searches: aspnet how to bind category labels to the radial axis of a radar chart with Aspose.Cells | c# create radar chart from worksheet data using Aspose.Cells | set custom axis labels on Aspose.Cells radar chart example | Aspose.Cells radar chart series range A2:A6,B2:B6 explanation
// Tags: Aspose.Cells radar chart with category axis labels | C# create radar chart from worksheet range | set radial axis labels Aspose.Cells | add series using combined range Aspose.Cells | save radar chart to Excel file Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills column A with category names and column B with numeric values, adds a radar chart, binds the series to the combined range A2:A6,B2:B6 so the radial axis displays those category names, sets a chart title, and saves the workbook as RadarChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data: categories in column A, values in column B
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["A5"].PutValue("Delta");
            sheet.Cells["A6"].PutValue("Epsilon");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(30);
            sheet.Cells["B6"].PutValue(25);

            // Add a radar chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series with categories and values in a single range
            // The range "A2:A6,B2:B6" sets category data (A) and value data (B)
            chart.NSeries.Add("A2:A6,B2:B6", true);

            // Optional: set a chart title
            chart.Title.Text = "Sample Radar Chart";

            // Save the workbook
            workbook.Save("RadarChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
