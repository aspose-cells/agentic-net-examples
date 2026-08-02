// Title: Clone Excel Workbook, Reassign Worksheet TabIds, Preserve Shapes – Aspose.Cells C# Example
// Description: Demonstrates how to copy a workbook with Aspose.Cells for .NET, assign new sequential TabId values to each worksheet in the cloned file, and retain all original shape collections (charts, images, drawings) before saving the result.
// Keywords: Aspose.Cells copy workbook C# | Worksheet TabId property | preserve shapes Aspose.Cells | clone Excel file .NET | duplicate workbook with drawings | reset worksheet TabId | Aspose.Cells workbook cloning
// Common Searches: how to clone a workbook and change TabId with Aspose.Cells | preserve charts and images when copying Excel file in .NET | assign new TabId to worksheets after workbook copy | Aspose.Cells copy workbook keep drawings | reset worksheet TabId after cloning
// Developer Intent: Create an exact copy of an existing workbook, give each sheet a fresh TabId, and keep all embedded shapes unchanged.
// Use Cases: Generate client‑specific reports from a master template while avoiding TabId conflicts. | Automate versioned workbook creation where sheet identifiers must be renumbered but charts and images stay intact. | Batch‑process multiple Excel files, duplicating them and normalizing TabIds for downstream systems.
// AI Prompts: Write C# code using Aspose.Cells to copy a workbook, assign sequential TabId values to each worksheet, and confirm shape counts match the source. | Explain the impact of Worksheet.TabId on Excel files and best practices for modifying it after cloning with Aspose.Cells. | Show how to clone a workbook, preserve all shape collections, and rename worksheets based on their new TabId values.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCloneExample
{
    // Demonstrates how to copy a workbook with Aspose.Cells for .NET, assign new sequential TabId values to each worksheet in the cloned file, and retain all original shape collections (charts, images, drawings) before saving the result.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Create a new empty workbook that will hold the clone
            Workbook clonedWorkbook = new Workbook();

            // Clone the entire workbook (includes worksheets, data, formats, and shapes)
            clonedWorkbook.Copy(sourceWorkbook);

            // Assign new TabId values to each worksheet in the cloned workbook
            int newTabId = 1;
            foreach (Worksheet ws in clonedWorkbook.Worksheets)
            {
                ws.TabId = newTabId++;
            }

            // At this point the shape collections are already preserved by the Copy method.
            // If you need to verify, you can compare the shape counts:
            // for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            // {
            //     Console.WriteLine($"Sheet {i}: Source shapes = {sourceWorkbook.Worksheets[i].Shapes.Count}, Cloned shapes = {clonedWorkbook.Worksheets[i].Shapes.Count}");
            // }

            // Save the cloned workbook
            clonedWorkbook.Save("cloned.xlsx");
        }
    }
}
