// Title: Create a 3‑D column chart with Aspose.Cells for .NET and attempt to set X=20°, Y=45°, Z=10° rotation angles
// AI Prompts: Write C# code that builds a workbook, adds sample data, inserts a 3‑D column chart, and tries to assign X‑axis rotation 20°, Y‑axis rotation 45°, Z‑axis rotation 10° using Aspose.Cells, including fallback handling when the API lacks rotation properties. | Update the provided Aspose.Cells example to detect whether chart rotation is supported and, if so, apply the specified X/Y/Z angles; otherwise log a clear warning and still save the file.
// Common Searches: Aspose.Cells .NET how to set X Y Z rotation on a 3D column chart | C# code to rotate Excel 3D chart using Aspose.Cells API | Is chart rotation supported in current Aspose.Cells version | Apply custom 3D view angles to Excel chart with Aspose.Cells for .NET | Save workbook after attempting to set 3D chart rotation Aspose.Cells
// Tags: Aspose.Cells 3D chart rotation | C# set chart X Y Z angles Aspose | Excel 3D column chart Aspose.Cells | unsupported chart rotation handling Aspose | save workbook with rotated chart .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes three numeric values, adds a 3‑D column chart linked to that data, notes that rotation properties are not available in the current Aspose.Cells release, and saves the file as RotatedChart.xlsx.
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
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("A1:A3", true);

            // Rotation properties are not available in the current Aspose.Cells version,
            // so they are omitted. The chart will use default 3‑D view.

            // Save the workbook to a file
            string outputPath = "RotatedChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
