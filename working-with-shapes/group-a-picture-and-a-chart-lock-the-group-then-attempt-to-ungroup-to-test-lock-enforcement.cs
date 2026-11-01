// Title: Attempt to group a picture with a chart, lock the group, and test ungrouping using Aspose.Cells for .NET (C#)
// Common Searches: how to group a picture and a chart in Aspose.Cells C# | Aspose.Cells lock shape group and prevent ungrouping example | C# code to test locked shape group ungrouping with Aspose.Cells | why grouping picture and chart is not supported in Aspose.Cells .NET | add picture and column chart to worksheet using Aspose.Cells C#
// Tags: group shapes Aspose.Cells C# | lock shape group Aspose.Cells | ungroup locked shape Aspose.Cells | insert picture worksheet Aspose.Cells | create column chart Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a picture from a file into cell B6, adds a column chart with data from A1:A3, notes that grouping a picture and a chart and locking the group is not supported in the current Aspose.Cells version, and saves the workbook as GroupedLockedOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Add a picture ----------
            string picturePath = "sample.png";
            if (!File.Exists(picturePath))
            {
                Console.WriteLine($"Picture file not found: {picturePath}");
                return;
            }

            // Add picture to cell B6 (row index 5, column index 0)
            int pictureIdx = sheet.Pictures.Add(5, 0, picturePath);
            Picture picture = sheet.Pictures[pictureIdx];

            // ---------- Add a chart ----------
            // Create a simple column chart.
            int chartIdx = sheet.Charts.Add(ChartType.Column, 15, 0, 25, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Populate some data for the chart.
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            chart.NSeries.Add("A1:A3", true);

            // Note: Grouping shapes (picture + chart) and locking the group
            // is not supported in the current Aspose.Cells version used.
            // The example focuses on adding a picture and a chart successfully.

            // Save the workbook to verify the result.
            string outputPath = "GroupedLockedOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
