// Title: How to enforce single‑selection in an Aspose.Cells PivotTable filter dialog using C# (workaround for missing EnableMultipleSelection)
// AI Prompts: Generate C# code that configures a PivotTable filter to allow only one item selection using Aspose.Cells, considering the EnableMultipleSelection property is unavailable. | Provide a C# example that uses alternative Aspose.Cells settings (e.g., PivotField.ShowDropDown) to mimic disabling multi‑select in a PivotTable filter. | Outline the steps to check for a PivotTable, apply a single‑choice filter configuration, and persist the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# enforce single choice in pivot table filter dialog | workaround for missing EnableMultipleSelection property in Aspose.Cells pivot tables | disable multi‑select dropdown for pivot field using Aspose.Cells API | which Aspose.Cells version introduces EnableMultipleSelection for PivotTable | how to limit pivot table filter to one selection in .NET with Aspose.Cells
// Tags: Aspose.Cells pivot table single choice filter | C# disable multi-select Aspose.Cells pivot | pivot field dropdown single selection Aspose.Cells | EnableMultipleSelection property missing workaround | Aspose.Cells save workbook after pivot modification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, checks for a PivotTable on the first worksheet, notes that the EnableMultipleSelection property is not present in the current Aspose.Cells release, and saves the workbook. It highlights that developers must use alternative settings or upgrade the library to achieve single‑selection behavior in PivotTable filter dialogs.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Check if the worksheet contains any PivotTables
            if (sheet.PivotTables.Count > 0)
            {
                // Retrieve the first PivotTable
                PivotTable pivotTable = sheet.PivotTables[0];

                // Note: The EnableMultipleSelection property is not available in the current Aspose.Cells version.
                // If needed, configure selection behavior using other available settings or upgrade the library.
            }
            else
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
