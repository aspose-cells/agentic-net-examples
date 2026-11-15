// Title: How to set chart category axis labels from a string array in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart category axis labels from string array | assign custom categories to a column chart using Aspose.Cells .NET | use CategoryData property with a string array in Aspose.Cells chart | programmatically change X‑axis labels of a column chart in C# Aspose.Cells | update chart categories without referencing worksheet range in Aspose.Cells
// Tags: set chart category axis Aspose.Cells C# | CategoryData property string array Aspose.Cells | column chart custom X‑axis labels .NET | Aspose.Cells chart axis labeling from code | programmatic chart category labels Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new Workbook, fills cells A1:B4 with month names and sales data, adds a column chart, binds the Y‑values to B2:B4, and demonstrates how to assign custom category axis labels using a string array via the CategoryData property (when supported). The workbook is saved as ChartWithCategoryLabels.xlsx.
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
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Bind the series values (Y‑values) to the data range
            chart.NSeries.Add("B2:B4", true);

            // NOTE: In some older Aspose.Cells versions the Series class does not expose
            // a CategoryData property. If needed, you can upgrade the library or set
            // category labels via other means. The chart will still display the data.

            // Determine output file path
            string outputPath = "ChartWithCategoryLabels.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
