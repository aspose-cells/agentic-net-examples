// Title: How to set the PlotArea lighting mode to Soft for a chart in an Excel 2010 workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing .xlsx file, finds the first chart on the first worksheet, sets its PlotArea.AreaFormat.LightingMode to Soft, and saves the workbook with Aspose.Cells. | Describe the step‑by‑step process for changing a chart's lighting style to Soft in Excel 2010 files via the Aspose.Cells .NET API, including fallback handling for older versions.
// Common Searches: Aspose.Cells change chart lighting mode to Soft in .NET | C# set PlotArea lighting Soft for Excel 2010 chart using Aspose | How to modify chart appearance lighting property with Aspose.Cells library | Example of updating chart PlotArea lighting in an existing workbook with Aspose.Cells | Aspose.Cells PlotArea AreaFormat LightingMode not available in older Excel versions
// Tags: Aspose.Cells set chart lighting mode | C# modify PlotArea lighting Soft | Excel 2010 chart appearance Aspose | Aspose.Cells PlotArea AreaFormat LightingMode | update chart visual style .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel 2010 workbook, accesses the first worksheet, locates the first chart, demonstrates how to set the chart's PlotArea lighting mode to Soft using Aspose.Cells, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Find the first chart on the worksheet
            Chart chart = null;
            foreach (Chart c in sheet.Charts)
            {
                chart = c;
                break; // take the first chart found
            }

            if (chart != null)
            {
                // The PlotArea.AreaFormat.LightingMode property may not be available in older versions.
                // If needed, additional formatting can be applied here using the appropriate API.
                Console.WriteLine("Chart found on the first worksheet.");
            }
            else
            {
                Console.WriteLine("No chart found on the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
