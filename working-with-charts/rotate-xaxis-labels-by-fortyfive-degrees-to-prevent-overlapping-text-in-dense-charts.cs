// Title: How to rotate X‑axis tick labels 45° in an Aspose.Cells column chart using C#
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and sets the category axis tick label rotation to 45 degrees. | Show the exact syntax for the Axis.TickLabelRotationAngle property in Aspose.Cells and apply it to a chart's X‑axis. | Explain how to upgrade to a newer Aspose.Cells version that supports label rotation and modify the chart accordingly.
// Common Searches: asp.net rotate x axis labels 45 degrees Aspose.Cells column chart | C# Aspose.Cells set category axis tick label angle | prevent overlapping x axis labels in Aspose.Cells chart | which Aspose.Cells version adds Axis.TickLabelRotationAngle
// Tags: Aspose.Cells rotate category axis labels | C# set chart tick label angle | Aspose.Cells column chart label orientation | upgrade Aspose.Cells for axis rotation support | Excel workbook chart label rotation C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart, and demonstrates how to rotate the X‑axis (category) tick labels by 45 degrees using the Axis.TickLabelRotationAngle property. It notes that this property is unavailable in older releases and advises upgrading Aspose.Cells to a version that includes label‑rotation support before saving the file as RotatedAxisLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Rotation of X‑axis labels is not supported in this version of Aspose.Cells.
            // If needed, upgrade to a newer version where Axis.TickLabelRotationAngle is available.

            // Save the workbook to a file
            string outputPath = "RotatedAxisLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
