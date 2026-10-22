// Title: Add data labels to a column chart and freeze the header row in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook containing a column chart that annotates each column with its numeric value and its category label, then lock the top row using FreezePanes, all via Aspose.Cells in C#. | Write C# code that builds a worksheet, inserts sample data, creates a column chart, turns on ShowValue and ShowCategoryName for the series, and freezes the first row with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to display category names on column chart data labels | C# freeze top row in Excel workbook using Aspose.Cells | show value and category name in chart series Aspose.Cells .NET example | add data labels to chart and keep header visible with FreezePanes Aspose.Cells
// Tags: Aspose.Cells chart data labels ShowValue ShowCategoryName | Aspose.Cells FreezePanes header row | Aspose.Cells create column chart C# | Aspose.Cells add data labels Excel chart | Aspose.Cells freeze first worksheet row

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data, inserts a column chart, enables data labels to show both values and category names, freezes the first worksheet row to keep the labels visible, and saves the file as ChartWithDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels: show values and category names
            chart.NSeries[0].DataLabels.ShowValue = true;            // display the numeric value
            chart.NSeries[0].DataLabels.ShowCategoryName = true;    // display the category name

            // Freeze the first row to keep it always visible
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook to a file
            string outputPath = "ChartWithDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
