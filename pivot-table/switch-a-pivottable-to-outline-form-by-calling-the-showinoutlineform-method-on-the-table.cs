// Title: Switch a PivotTable to Outline form using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file, locates the first PivotTable, calls ShowInOutlineForm(), and saves the workbook to a new file. | Demonstrate how to use Aspose.Cells to change a PivotTable's display mode to Outline view while handling missing input files. | Provide a step‑by‑step example of enabling outline form for a PivotTable in a .NET console application with error handling.
// Common Searches: Aspose.Cells C# example to enable outline view for a PivotTable | How to programmatically set PivotTable ShowInOutlineForm in .NET | C# load workbook, modify PivotTable layout to outline, save file using Aspose.Cells | Switch PivotTable to outline form with Aspose.Cells for .NET tutorial
// Tags: Aspose.Cells PivotTable ShowInOutlineForm | C# modify PivotTable layout outline | Aspose.Cells load workbook adjust PivotTable | save Excel workbook after PivotTable changes Aspose.Cells | error handling missing Excel file Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// // Loads an existing Excel workbook, retrieves the first PivotTable on the first worksheet, switches it to Outline form by calling ShowInOutlineForm(), and saves the modified workbook to a new file, with basic file‑existence checks and error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Switch the PivotTable to Outline form (method, not property)
            pivotTable.ShowInOutlineForm();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
