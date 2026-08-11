// Title: Clone an Excel workbook, preserve shapes, and assign new TabId values with Aspose.Cells for .NET (C#)
// Description: Load a source workbook, create an empty workbook, copy all worksheets, manually duplicate each shape to keep drawings, assign sequential TabId numbers (starting at 1000) to every cloned sheet, and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# workbook clone | copy shapes Aspose.Cells | Worksheet TabId | preserve drawings Excel | Workbook.Copy | clone workbook with shapes | assign TabId programmatically
// Common Searches: Aspose.Cells clone workbook with shapes C# | How to copy worksheets and keep drawings Aspose.Cells | Set custom TabId for worksheets after cloning | Preserve shape collections when copying Excel file using Aspose.Cells | C# assign sequential TabId to cloned sheets
// Developer Intent: Duplicate an existing Excel file, retain all embedded shapes, and give each copied worksheet a unique TabId.
// Use Cases: Generate client‑specific reports by cloning a master template, keeping charts and images intact, and assigning distinct TabId values to avoid identifier conflicts. | Automate versioned copies of a financial model where each copy preserves drawings and receives a new TabId range for change tracking. | Create multiple workbooks from a single source for batch processing, ensuring embedded objects remain and each sheet has a sequential TabId for downstream APIs.
// AI Prompts: Provide C# code using Aspose.Cells to clone a workbook, copy all shapes, and set new TabId values for each worksheet. | Show an example of copying worksheets from one Excel file to another while preserving drawings and assigning sequential TabId numbers starting at a custom base. | Explain how Aspose.Cells handles shape duplication and TabId assignment when using Workbook.Copy in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCloneExample
{
    // Load a source workbook, create an empty workbook, copy all worksheets, manually duplicate each shape to keep drawings, assign sequential TabId numbers (starting at 1000) to every cloned sheet, and save the result using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";
                const string outputPath = "ClonedWorkbook.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new empty workbook for the clone
                Workbook clonedWorkbook = new Workbook();

                // Copy all contents (worksheets, data, formulas, styles, etc.)
                clonedWorkbook.Copy(sourceWorkbook);

                // Preserve original shape collections by copying each shape manually
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    Worksheet sourceSheet = sourceWorkbook.Worksheets[i];
                    Worksheet destSheet = clonedWorkbook.Worksheets[i];

                    foreach (Shape sourceShape in sourceSheet.Shapes)
                    {
                        // Copy shape using its bounding rows and columns
                        destSheet.Shapes.AddCopy(
                            sourceShape,
                            sourceShape.UpperLeftRow,
                            sourceShape.UpperLeftColumn,
                            sourceShape.LowerRightRow,
                            sourceShape.LowerRightColumn);
                    }
                }

                // Assign new TabId values to all sheets in the cloned workbook
                const int tabIdBase = 1000;
                for (int i = 0; i < clonedWorkbook.Worksheets.Count; i++)
                {
                    clonedWorkbook.Worksheets[i].TabId = tabIdBase + i;
                }

                // Save the cloned workbook
                clonedWorkbook.Save(outputPath);
                Console.WriteLine($"Cloned workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
