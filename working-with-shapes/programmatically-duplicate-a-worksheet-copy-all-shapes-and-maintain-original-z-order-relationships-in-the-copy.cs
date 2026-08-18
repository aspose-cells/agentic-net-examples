// Title: Copy a worksheet with all shapes and preserve Z‑order using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, duplicates a worksheet with AddCopy, iterates through the source sheet's Shapes collection, adds each shape to the new sheet at the same cell coordinates, and copies the original ZOrderPosition to keep the stacking order before saving the file.
// Keywords: Aspose.Cells | C# | .NET | duplicate worksheet | copy worksheet with shapes | preserve Z‑order | AddCopy | Shape ZOrderPosition | Excel drawing objects | Workbook example
// Common Searches: Aspose.Cells copy worksheet with shapes | preserve shape Z‑order when duplicating Excel sheet | C# duplicate sheet and keep drawing order Aspose.Cells | AddCopy shapes ZOrderPosition example | how to clone a worksheet with charts and images in Aspose.Cells
// Developer Intent: Create an exact copy of a worksheet that includes every shape and retains the original layering order.
// Use Cases: Generate client‑specific reports by cloning a template sheet that contains positioned charts and logos. | Automate monthly dashboards where the visual hierarchy of graphics must stay unchanged across copies. | Back up complex worksheets with layered images, text boxes, and charts without losing their Z‑order.
// AI Prompts: Show C# code to duplicate an Excel worksheet with all shapes while keeping Z‑order using Aspose.Cells. | Explain how to copy shapes and preserve their ZOrderPosition after worksheet duplication in Aspose.Cells for .NET. | Provide a step‑by‑step guide to verify that shape stacking order is identical in the original and copied sheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads a workbook, duplicates a worksheet with AddCopy, iterates through the source sheet's Shapes collection, adds each shape to the new sheet at the same cell coordinates, and copies the original ZOrderPosition to keep the stacking order before saving the file.
    public class WorksheetDuplicateWithShapes
    {
        // Entry point required by the project
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string outputPath = "WorkbookWithDuplicatedSheet.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the worksheet to be duplicated
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Duplicate the worksheet (copies cells, formats, and drawing objects)
            int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = sourceSheet.Name + "_Copy";

            // Copy each shape preserving its Z‑order
            foreach (Shape srcShape in sourceSheet.Shapes)
            {
                // Add a copy of the shape to the new worksheet at the same position
                Shape newShape = copiedSheet.Shapes.AddCopy(
                    srcShape,
                    srcShape.UpperLeftRow,
                    srcShape.UpperLeftColumn,
                    srcShape.Height,
                    srcShape.Width);

                // Preserve the original Z‑order position
                newShape.ZOrderPosition = srcShape.ZOrderPosition;
            }

            // Save the workbook with the duplicated sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
