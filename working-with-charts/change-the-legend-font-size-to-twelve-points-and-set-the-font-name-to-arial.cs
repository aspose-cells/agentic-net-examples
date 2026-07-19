// Title: C# – Apply Arial 12 pt font to a chart legend with Aspose.Cells
// Description: The sample builds a workbook, inserts a column chart, and programmatically changes the legend’s typeface to Arial and its size to 12 points using Aspose.Cells for .NET, then writes the file as LegendFontExample.xlsx.
// Keywords: Aspose.Cells C# | chart legend font | Arial font Excel | 12 point legend | modify chart legend Aspose | column chart styling | Excel automation Aspose.Cells | set legend font size
// Common Searches: How to change chart legend font in Aspose.Cells C# | Set legend to Arial 12pt in Excel using Aspose.Cells | Aspose.Cells chart legend formatting example | C# code to modify legend appearance in an Excel workbook
// Developer Intent: Programmatically set the legend’s typeface to Arial and its size to 12 pt in an Excel chart generated with Aspose.Cells.
// Use Cases: Enforce corporate branding by standardizing legend fonts across all generated reports. | Create presentation‑ready Excel files where the legend matches the visual style of slide decks. | Improve readability of dashboards exported to Excel for users on high‑resolution displays.
// AI Prompts: Generate C# code with Aspose.Cells that changes a chart legend to Times New Roman, 10 pt, for a pie chart. | Show how to apply the same Arial 12 pt legend font to every chart in a workbook using Aspose.Cells. | Provide an example that sets legend font color, name, and size for a line chart in Aspose.Cells (C#).

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample builds a workbook, inserts a column chart, and programmatically changes the legend’s typeface to Arial and its size to 12 points using Aspose.Cells for .NET, then writes the file as LegendFontExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Change the legend font: set name to Arial and size to 12 points
        chart.Legend.Font.Name = "Arial";
        chart.Legend.Font.Size = 12;

        // Save the workbook to a file
        workbook.Save("LegendFontExample.xlsx");
    }
}
