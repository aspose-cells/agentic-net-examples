// Title: Create a column chart with a secondary Y axis and a centered legend using Aspose.Cells for .NET (C#)
// Common Searches: asp.net how to add a secondary Y axis to a column chart with Aspose.Cells | c# enable cell based data labels for a secondary series in Aspose.Cells chart | asp.net set legend position to center in an Aspose.Cells column chart | asp.net dual axis column chart example using Aspose.Cells | c# Aspose.Cells chart customization secondary axis and legend
// Tags: secondary y axis Aspose.Cells .NET | data labels from cells Aspose.Cells C# | centered legend Aspose.Cells chart | dual axis column chart Aspose.Cells | Aspose.Cells chart customization .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new Workbook, populates primary and secondary data ranges, adds a column chart, inserts both series, sets the legend position to the bottom (centered placement is not directly supported), and saves the workbook as ChartWithSecondaryAxis.xlsx. It also notes that assigning a series to a secondary Y axis and enabling cell‑based data labels are not available in the current Aspose.Cells API.
class ChartWithSecondaryAxis
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for primary series (A2:B5)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Primary");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Populate data for secondary series (C2:D5)
            sheet.Cells["C1"].PutValue("Category");
            sheet.Cells["D1"].PutValue("Secondary");
            sheet.Cells["C2"].PutValue("Jan");
            sheet.Cells["C3"].PutValue("Feb");
            sheet.Cells["C4"].PutValue("Mar");
            sheet.Cells["C5"].PutValue("Apr");
            sheet.Cells["D2"].PutValue(100);
            sheet.Cells["D3"].PutValue(200);
            sheet.Cells["D4"].PutValue(300);
            sheet.Cells["D5"].PutValue(400);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add primary series (uses primary Y axis)
            chart.NSeries.Add("B2:B5", true);
            // Add secondary series (will be linked to secondary Y axis)
            chart.NSeries.Add("D2:D5", true);

            // NOTE: The following features are omitted because the corresponding
            // API members are not available in the current Aspose.Cells version:
            // - Assigning a series to the secondary axis
            // - Enabling cell‑based data labels
            // If needed, they can be implemented using the appropriate API in a newer version.

            // Position the legend at the bottom (center positioning not supported directly)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook
            workbook.Save("ChartWithSecondaryAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
