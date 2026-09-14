// Title: Create a column chart from cell data and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, fills cells A1:B4 with category and value data, adds a column chart positioned from row 5 to 20 and column A to H, binds the series to B2:B4 and categories to A2:A4, sets the chart title, and saves the file as ReflectionDemo.xlsx with Aspose.Cells. | Generate a complete Aspose.Cells example that demonstrates how to programmatically insert a column chart into the first worksheet, configure its data source and title, and export the workbook to an .xlsx file.
// Common Searches: Aspose.Cells C# add column chart to worksheet from cell range | how to bind series and category data for a chart using Aspose.Cells .NET | save Excel workbook with chart using Aspose.Cells C# example | set chart title programmatically Aspose.Cells .NET
// Tags: Aspose.Cells add column chart C# | Aspose.Cells bind chart data range .NET | Aspose.Cells set chart title C# | Aspose.Cells save workbook with chart .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new Workbook, populates cells A1:B4 with sample categories and values, inserts a column chart positioned at rows 5‑20 and columns A‑H, links the series to B2:B4 and categories to A2:A4, assigns a title, and saves the workbook as ReflectionDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set chart title (reflection not supported in this version)
            chart.Title.Text = "Sample Column Chart";

            // Output confirmation
            Console.WriteLine($"Chart title set: {chart.Title.Text}");

            // Save the workbook
            workbook.Save("ReflectionDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
