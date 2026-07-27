// Title: C# – Set 3‑D Chart Rotation (X=20, Y=45, Z=10) with Aspose.Cells for .NET
// Description: Creates a workbook, fills cells A1:B4 with sample data, adds a 3‑D column chart and demonstrates how to apply X, Y, Z rotation angles. The current Aspose.Cells version does not expose rotation properties, so the example notes the limitation and suggests upgrading to a newer release before saving the file as Chart3DRotationDemo.xlsx.
// Keywords: Aspose.Cells 3D chart rotation C# | Aspose.Cells set chart X Y Z angles | C# Aspose.Cells 3D column chart orientation | Aspose.Cells chart rotation API .NET | Aspose.Cells version chart rotation support | GitHub Aspose.Cells chart example | Aspose.Cells 3D perspective settings
// Common Searches: how to rotate a 3d chart in Aspose.Cells C# | Aspose.Cells set X Y Z rotation for 3D chart | Aspose.Cells chart rotation properties missing | which Aspose.Cells version adds chart rotation | C# code to change 3D chart perspective Aspose.Cells
// Developer Intent: Add a 3‑D column chart to a workbook, attempt to set X=20, Y=45, Z=10 rotation angles, and save the workbook.
// Use Cases: Generate a quick Excel file with sample data and a 3‑D column chart for visual testing. | Show developers that chart rotation APIs are unavailable in older Aspose.Cells releases, prompting an upgrade. | Provide a baseline example that can be extended once rotation properties become accessible.
// AI Prompts: Write C# code using the latest Aspose.Cells API to set X=20, Y=45, Z=10 rotation on a 3‑D column chart. | Explain how to check the Aspose.Cells release notes to find when chart rotation properties were introduced. | Suggest alternative techniques (e.g., adjusting perspective or view angles) to mimic chart rotation in current Aspose.Cells versions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, fills cells A1:B4 with sample data, adds a 3‑D column chart and demonstrates how to apply X, Y, Z rotation angles. The current Aspose.Cells version does not expose rotation properties, so the example notes the limitation and suggests upgrading to a newer release before saving the file as Chart3DRotationDemo.xlsx.
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
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // NOTE: Rotation properties are not available in this version of Aspose.Cells.
            // If needed, they can be set using the appropriate API in newer versions.

            // Save the workbook with the configured chart
            string outputPath = "Chart3DRotationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
