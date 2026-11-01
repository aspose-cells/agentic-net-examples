// Title: Configure Z‑axis tick interval to display every second depth level on a 3D clustered column chart using Aspose.Cells for .NET (currently unsupported)
// Common Searches: Aspose.Cells C# set Z axis tick interval for 3D column chart | display every other depth level on a 3D clustered column chart Aspose.Cells | how to adjust Z axis spacing in 3D Excel chart with Aspose.Cells .NET | Aspose.Cells 3D chart depth level labeling options | Z axis interval not available in current Aspose.Cells version
// Tags: Aspose.Cells 3D chart Z axis interval | C# configure Z axis tick spacing Aspose.Cells | 3D clustered column chart depth level labeling .NET | Excel 3D chart axis configuration Aspose.Cells | Aspose.Cells chart axis settings limitation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a 3D clustered column chart, and demonstrates that the Z‑axis interval (to show every second depth level) cannot be set with the current Aspose.Cells API, then saves the file as 3DChart_ZAxisInterval.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a 3D clustered column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3DClustered, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: The ZAxis property is not available in the current Aspose.Cells version.
            // If needed, configure Z‑axis settings using the appropriate API when it becomes available.

            // Save the workbook to a file
            string outputPath = "3DChart_ZAxisInterval.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
