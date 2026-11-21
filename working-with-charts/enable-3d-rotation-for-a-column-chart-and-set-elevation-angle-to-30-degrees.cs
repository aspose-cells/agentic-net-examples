// Title: Generate a 3‑D Column chart in Aspose.Cells for .NET (rotation and elevation not available)
// Common Searches: how to create a 3d column chart with Aspose.Cells in C# | Aspose.Cells column3d chart rotation properties C# | set elevation angle for Excel 3d chart using Aspose.Cells .NET | C# Aspose.Cells 3d chart view angle limitation | add 3d column chart to workbook with Aspose.Cells
// Tags: Aspose.Cells create 3d column chart | Aspose.Cells column3d rotation unsupported | Aspose.Cells chart elevation angle limitation | Aspose.Cells .NET Excel 3d visualization | Aspose.Cells chart view angle configuration

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, populates sample data, adds a 3‑D column chart (ChartType.Column3D), defines the series and category ranges, and saves the file. It also notes that the RotationX, RotationY, and Elevation properties are not available in the current Aspose.Cells version, so 3‑D rotation and elevation angle cannot be set programmatically.
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
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a 3‑D column chart (type Column3D)
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Note: RotationX, RotationY, and Elevation properties are not available
            // in the current Aspose.Cells version; they have been omitted.

            // Define output file path
            string outputPath = "3DColumnChart.xlsx";

            // Delete existing file to avoid conflicts
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
