// Title: Move an Aspose.Cells column chart to start at row 5, column 3 while keeping its size (C#)
// AI Prompts: Use the Chart.Move method to relocate a column chart so its upper‑left corner is at row 5, column 3 without altering its height or width. | Set the UpperLeftRow and UpperLeftColumn of a chart's Position in C# to row 5 and column 3 while preserving the original chart dimensions.
// Common Searches: Aspose.Cells C# move chart to specific cell coordinates row 5 column 3 | how to set chart upper left cell in an Aspose.Cells workbook | preserve chart size when repositioning a chart with Aspose.Cells | Chart.Move method parameters explanation Aspose.Cells C# | set chart location by row and column using Aspose.Cells API
// Tags: Aspose.Cells Chart.Move method C# | chart upper-left cell positioning Aspose.Cells | preserve chart size Aspose.Cells | Excel chart relocation using Aspose.Cells | C# set chart position by row and column

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartPositionDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, then moves the chart so its top‑left corner is at row 5, column 3 while retaining its original size, and finally saves the file as ChartPositionDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart. Initial position: rows 5‑15, columns 0‑5
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Move the chart so that its top‑left corner is at row 5, column 3
            // Keep the same size (rows 5‑15, columns 3‑8)
            chart.Move(5, 3, 15, 8);

            // Optional: display the new position in the console
            Console.WriteLine($"Chart new top row: {chart.ChartObject.UpperLeftRow}");
            Console.WriteLine($"Chart new left column: {chart.ChartObject.UpperLeftColumn}");

            // Save the workbook
            workbook.Save("ChartPositionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
