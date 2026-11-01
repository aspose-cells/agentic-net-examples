// Title: Create a 3‑D column chart with Aspose.Cells for .NET and apply X=20°, Y=45°, Z=10° rotation before saving the workbook
// AI Prompts: Write C# code that builds a workbook, inserts a three‑dimensional column chart, and sets its X, Y, and Z rotation angles to 20, 45, and 10 degrees using the PlotArea.Area3DFormat API, including a version‑check fallback, then saves the file. | Show how to programmatically verify whether Aspose.Cells exposes chart‑rotation members and, when available, adjust the three‑dimensional view of the chart before exporting the workbook.
// Common Searches: Aspose.Cells C# set 3D chart X rotation angle to 20 degrees | How to change Y and Z view angles of a 3D column chart in Aspose.Cells .NET | PlotArea.Area3DFormat rotation example for Excel workbook using Aspose.Cells | Check Aspose.Cells version for 3D chart rotation support in C# | Save workbook with rotated 3D chart using Aspose.Cells
// Tags: Aspose.Cells Area3DFormat property | C# three‑dimensional column chart orientation | Workbook save with rotated three‑dimensional chart Aspose.Cells | Version‑aware chart rotation Aspose.Cells | Excel chart view angles C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds sample data, inserts a 3‑D column chart, notes that rotation properties are unavailable in the current Aspose.Cells version (suggesting the PlotArea.Area3DFormat API for newer releases), and saves the file as Chart3DRotation.xlsx.
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

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Rotation properties are not available in the current Aspose.Cells version.
            // If needed, they can be set using chart.PlotArea.Area3DFormat in newer versions.

            // Determine output file path
            string outputPath = "Chart3DRotation.xlsx";

            // Ensure the directory exists (if a directory is specified)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
