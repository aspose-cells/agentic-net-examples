// Title: How to change the second series to a line type in a combo chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook with sample data, add a column combo chart, then set the second series to ChartType.Line using Aspose.Cells in C#. | Programmatically modify an existing Aspose.Cells chart to change the second data series from column to line after the chart has been created. | Generate a mixed column‑line combo chart in C# by adding two series and converting the second series to a line series post‑creation.
// Common Searches: Aspose.Cells C# change series type to line in combo chart after adding series | set second series as line in mixed column line chart using Aspose.Cells .NET | how to convert a column series to line in Aspose.Cells chart programmatically
// Tags: Aspose.Cells change series type programmatically | combo chart column to line Aspose.Cells | C# Aspose.Cells mixed chart series conversion | modify chart series type after creation .NET | Aspose.Cells chart customization line series

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data, adds a combo chart initially as a column chart with two series, then changes the second series' type to a line chart, and saves the workbook as ComboChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Series1");
            ws.Cells["C1"].PutValue("Series2");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["C2"].PutValue(15);
            ws.Cells["C3"].PutValue(25);
            ws.Cells["C4"].PutValue(35);

            // Add a combo chart (initially a column chart)
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = ws.Charts[chartIndex];

            // Add first series (column)
            chart.NSeries.Add("B2:B4", true);
            // Add second series (column initially)
            chart.NSeries.Add("C2:C4", true);

            // Change the second series type from column to line
            chart.NSeries[1].Type = ChartType.Line;

            // Define output file path
            string outputPath = "ComboChart.xlsx";

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
