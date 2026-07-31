// Title: Aspose.Cells for .NET – Set Chart Legend Font to Arial 12pt (C#)
// Description: C# code that creates a workbook, adds a column chart, and changes the legend’s font to Arial, 12 points using Aspose.Cells for .NET, then saves the file.
// Keywords: Aspose.Cells C# chart legend font | set legend font size Aspose.Cells | Arial 12pt legend Aspose.Cells | modify chart legend appearance .NET | Excel legend styling C# | Aspose.Cells chart formatting
// Common Searches: Aspose.Cells change chart legend font size | C# set legend font to Arial in Excel chart | How to format legend in Aspose.Cells chart | Set legend font properties programmatically Aspose.Cells | Example of customizing chart legend with Aspose.Cells .NET
// Developer Intent: Programmatically set a chart legend’s font name to Arial and its size to 12 points in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enforce corporate branding by applying a uniform legend font across all generated charts. | Enhance readability of exported Excel reports for end‑users by customizing legend typography. | Create a reusable helper method that formats legend appearance for any chart type during automated workbook creation.
// AI Prompts: Show me C# code to change the legend font name and size for any Aspose.Cells chart. | Provide a reusable function that accepts a Chart object and sets its legend to Arial 12pt. | Explain how to update legend font properties after loading an existing workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendFontDemo
{
    // C# code that creates a workbook, adds a column chart, and changes the legend’s font to Arial, 12 points using Aspose.Cells for .NET, then saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Change the legend font: set name to Arial and size to 12 points
            chart.Legend.Font.Name = "Arial";
            chart.Legend.Font.Size = 12;

            // Save the workbook
            workbook.Save("LegendFontExample.xlsx");
        }
    }
}
