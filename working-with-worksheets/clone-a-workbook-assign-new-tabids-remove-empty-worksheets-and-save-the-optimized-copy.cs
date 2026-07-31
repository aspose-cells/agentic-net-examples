// Title: Clone Workbook, Assign Unique TabIds, Delete Blank Sheets, and Save Optimized Excel with Aspose.Cells C#
// Description: Loads a source XLSX, creates a new Workbook, copies all worksheets, assigns a distinct TabId to each sheet, removes worksheets that contain no data, cleans unused styles, and saves the streamlined file as an optimized XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells clone workbook C# | set TabId worksheet | remove empty worksheets Aspose | optimize Excel file size .NET | remove unused styles Aspose.Cells | Workbook.Copy example | Excel cleanup automation
// Common Searches: How to copy a workbook and set new TabIds with Aspose.Cells | Delete blank worksheets after cloning an Excel file in C# | Reduce Excel file size by removing empty sheets and unused styles | Aspose.Cells example for workbook optimization
// Developer Intent: Generate a lightweight copy of an existing workbook, give each sheet a unique TabId, purge blank worksheets and unused styles, then save the cleaned file.
// Use Cases: Create distribution‑ready templates by cloning a master workbook, stripping placeholder sheets, and assigning unique TabIds to prevent tab conflicts. | Produce per‑user reports where each copy receives its own TabIds while eliminating empty worksheets to keep files compact. | Automate batch cleanup of generated Excel files—remove blank sheets, unused styles, and other redundancies before archiving or publishing.
// AI Prompts: Write C# code using Aspose.Cells that clones a workbook, assigns a unique TabId to every worksheet, removes sheets with no data, cleans unused styles, and saves the optimized workbook. | Explain why MaxDataRow and MaxDataColumn are reliable for detecting empty worksheets in Aspose.Cells and why backward iteration is required when deleting sheets. | Suggest further size‑reduction techniques for an Aspose.Cells workbook, such as removing unused named ranges, clearing empty columns, or compressing images.

using System;
using Aspose.Cells;

namespace WorkbookOptimizationDemo
{
    // Loads a source XLSX, creates a new Workbook, copies all worksheets, assigns a distinct TabId to each sheet, removes worksheets that contain no data, cleans unused styles, and saves the streamlined file as an optimized XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Load the source workbook (uses the Workbook(string) constructor)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty workbook (uses the parameterless Workbook() constructor)
            Workbook optimizedWorkbook = new Workbook();

            // Clone the source workbook into the new workbook (uses Workbook.Copy(Workbook))
            optimizedWorkbook.Copy(sourceWorkbook);

            // Assign new unique TabIds to each worksheet
            for (int i = 0; i < optimizedWorkbook.Worksheets.Count; i++)
            {
                // Example: set TabId to (i + 1) * 10 to ensure uniqueness
                optimizedWorkbook.Worksheets[i].TabId = (i + 1) * 10;
            }

            // Remove empty worksheets (iterate backwards to avoid index shift when removing)
            for (int i = optimizedWorkbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = optimizedWorkbook.Worksheets[i];

                // A worksheet is considered empty when it has no used cells.
                // MaxDataRow/MaxDataColumn return -1 when there is no data.
                bool isEmpty = sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0;

                if (isEmpty)
                {
                    optimizedWorkbook.Worksheets.RemoveAt(i);
                }
            }

            // Optional: clean up unused styles to reduce file size
            optimizedWorkbook.RemoveUnusedStyles();

            // Save the optimized workbook (uses Workbook.Save(string, SaveFormat))
            string outputPath = "optimized_output.xlsx";
            optimizedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Optimized workbook saved to: {outputPath}");
        }
    }
}
