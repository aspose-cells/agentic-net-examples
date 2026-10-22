// Title: Enable 3‑D rotation and set a 30° elevation angle for a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, adds sample data, inserts a 3‑D column chart, sets RotationAngle to 45°, and shows how to apply ElevationAngle = 30° when the property is available. | Show how to programmatically adjust both horizontal rotation and elevation view of a 3‑D column chart in Aspose.Cells for .NET, including fallback handling for versions that lack ElevationAngle. | Provide a complete Aspose.Cells example that builds a 3‑D column chart, customizes its perspective angles, and saves the file, with comments on API limitations.
// Common Searches: Aspose.Cells C# set 3D column chart rotation angle to 45 degrees | How to change elevation angle of a 3D chart with Aspose.Cells .NET | C# example adding a 3D column chart and adjusting view angles in Aspose.Cells | Aspose.Cells chart rotation and elevation properties not supported in current version | Create a 3D column chart with custom perspective using Aspose.Cells for .NET
// Tags: Aspose.Cells 3D chart rotation angle C# | Aspose.Cells set column chart elevation angle .NET | C# Aspose.Cells create 3D column chart | Aspose.Cells chart view customization .NET | Aspose.Cells version limitation elevation angle

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The program creates a workbook, fills it with sample data, adds a 3‑D column chart, sets the chart's horizontal RotationAngle to 45°, notes that ElevationAngle is unavailable in the current Aspose.Cells version (default elevation is used), and saves the workbook as ColumnChart3D.xlsx.
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
            Cells cells = sheet.Cells;

            // Populate sample data for the column chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add a 3‑D column chart to the worksheet
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 6, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable 3‑D rotation (horizontal rotation)
            chart.RotationAngle = 45; // horizontal rotation

            // Note: ElevationAngle property is not available in the current Aspose.Cells version.
            // The default elevation angle will be used.

            // Save the workbook to a file
            workbook.Save("ColumnChart3D.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
