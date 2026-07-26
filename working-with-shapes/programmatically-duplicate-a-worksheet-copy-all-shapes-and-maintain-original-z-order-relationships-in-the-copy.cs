// Title: Duplicate a Worksheet and Copy All Shapes While Preserving Z‑Order with Aspose.Cells for .NET
// Description: Loads a workbook, creates a copy of a worksheet, iterates the source shapes in Z‑order, uses Shapes.AddCopy to duplicate each shape at the same location, restores the original ZOrderPosition, and saves the workbook with the cloned sheet and its layered drawings.
// Keywords: Aspose.Cells | duplicate worksheet | copy shapes | preserve Z‑order | AddCopy | C# | Excel automation | .NET | clone sheet with drawings | ZOrderPosition
// Common Searches: Aspose.Cells duplicate worksheet with shapes | Copy Excel shapes preserving Z‑order C# | How to clone a sheet and keep drawing objects | Shapes.AddCopy example Aspose.Cells | Programmatically copy worksheet drawings .NET
// Developer Intent: The developer needs to programmatically clone an Excel worksheet and replicate every shape on the new sheet, keeping the original stacking (Z‑order) intact.
// Use Cases: Generate multiple report tabs that share identical chart and image layouts, ensuring overlays appear in the same order. | Create department‑specific dashboard copies where layered graphics must retain their visual hierarchy. | Automate template‑based worksheets that require exact shape positioning and stacking for consistent branding.
// AI Prompts: Provide C# code to duplicate an Excel worksheet and copy all its shapes while preserving Z‑order using Aspose.Cells. | Explain error‑handling strategies when copying shapes with Shapes.AddCopy in Aspose.Cells. | Show how to adjust ZOrderPosition after copying shapes to maintain the original stacking sequence.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads a workbook, creates a copy of a worksheet, iterates the source shapes in Z‑order, uses Shapes.AddCopy to duplicate each shape at the same location, restores the original ZOrderPosition, and saves the workbook with the cloned sheet and its layered drawings.
    public class WorksheetDuplicateWithShapes
    {
        // Entry point for the example
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(sourcePath);

                // Retrieve the worksheet to duplicate
                Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
                if (sourceSheet == null)
                {
                    Console.WriteLine("Worksheet 'Sheet1' not found in the source workbook.");
                    return;
                }

                // Duplicate the worksheet (cells, formats, etc. are copied, but drawing objects are not)
                int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = sourceSheet.Name + "_Copy";

                // Copy all shapes from the source sheet to the copied sheet while preserving Z‑order
                var sourceShapes = sourceSheet.Shapes
                    .Cast<Shape>()
                    .OrderBy(s => s.ZOrderPosition)
                    .ToList();

                foreach (Shape srcShape in sourceShapes)
                {
                    try
                    {
                        // Add a copy of the shape at the same position on the target sheet
                        Shape newShape = copiedSheet.Shapes.AddCopy(
                            srcShape,
                            srcShape.UpperLeftRow,
                            srcShape.UpperLeftColumn,
                            srcShape.Height,
                            srcShape.Width);

                        // Preserve the original Z‑order position
                        newShape.ZOrderPosition = srcShape.ZOrderPosition;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to copy shape '{srcShape.Name}': {ex.Message}");
                    }
                }

                // Save the workbook with the duplicated worksheet and shapes
                const string outputPath = "DuplicatedWorkbook.xlsx";
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
