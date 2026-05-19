using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string destPath = "WorkbookWithDuplicatedSheet.xlsx";

            // Ensure the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(sourcePath);

                // Get the worksheet to duplicate
                Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
                if (sourceSheet == null)
                {
                    Console.WriteLine("Worksheet 'Sheet1' not found in the source workbook.");
                    return;
                }

                // Duplicate the worksheet (cells and formats only)
                int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = sourceSheet.Name + "_Copy";

                // Copy all shapes preserving Z‑order
                ShapeCollection sourceShapes = sourceSheet.Shapes;
                ShapeCollection destShapes = copiedSheet.Shapes;

                // Order shapes by their ZOrderPosition (front to back)
                var orderedShapes = sourceShapes.Cast<Shape>()
                                                .OrderBy(s => s.ZOrderPosition)
                                                .ToList();

                foreach (Shape srcShape in orderedShapes)
                {
                    // Add a copy of the shape to the destination worksheet at the same position
                    Shape newShape = destShapes.AddCopy(
                        srcShape,
                        srcShape.UpperLeftRow,
                        srcShape.UpperLeftColumn,
                        srcShape.LowerRightRow,
                        srcShape.LowerRightColumn);

                    // Preserve the original Z‑order position
                    newShape.ZOrderPosition = srcShape.ZOrderPosition;
                }

                // Save the workbook with the duplicated sheet and copied shapes
                workbook.Save(destPath);
                Console.WriteLine($"Workbook saved to {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}