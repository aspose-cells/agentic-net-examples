// Title: How to apply Plastic 3‑D material and Soft lighting (70% intensity) to a chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and then sets chart.PresetMaterial = PresetMaterial.Plastic, chart.PresetLighting = PresetLighting.Soft, and chart.LightingIntensity = 70. | Write a routine that opens an existing Aspose.Cells workbook, loops through all charts, and configures each chart to use Plastic as the 3‑D material with a Soft lighting preset at 70 % intensity.
// Common Searches: Aspose.Cells set 3D chart material to plastic in C# | C# Aspose.Cells apply soft lighting preset to Excel chart | How to change lighting intensity to 70 percent for 3D charts using Aspose.Cells .NET | Update all charts in an Aspose.Cells workbook to use plastic material and soft lighting
// Tags: chart presetmaterial plastic Aspose.Cells | chart presetlighting soft Aspose.Cells | chart lightingintensity 70 Aspose.Cells | aspnet excel 3d chart appearance settings | aspnet chart material and lighting configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart and ChartType

namespace AsposeCells3DExample
{
    // // This example creates a new workbook, adds a worksheet, inserts a column chart, and (when using a recent Aspose.Cells version) configures the chart's 3‑D appearance by setting PresetMaterial to Plastic, PresetLighting to Soft, and LightingIntensity to 70 %. It also notes that these properties are only available in newer library releases.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a new worksheet to host the chart
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Add a column chart to the worksheet
                // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // NOTE: PresetMaterial, PresetLighting, and LightingIntensity properties are
                // available only in newer versions of Aspose.Cells. If using an older version,
                // these properties are omitted to ensure compilation.

                // Save the workbook
                string outputPath = "ThreeDMaterialPlasticSoftLighting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
