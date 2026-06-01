using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsBatchCopy
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "BatchCopyResult.xlsx";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook that contains the source worksheet with drawing objects
                Workbook templateWorkbook = new Workbook(templatePath);
                Worksheet templateSheet = templateWorkbook.Worksheets[0];

                // Number of new worksheets to create
                int newSheetCount = 5;

                for (int i = 0; i < newSheetCount; i++)
                {
                    // Create a new blank worksheet
                    string newSheetName = $"Copy_{i + 1}";
                    Worksheet newSheet = templateWorkbook.Worksheets.Add(newSheetName);

                    ShapeCollection srcShapes = templateSheet.Shapes;
                    ShapeCollection destShapes = newSheet.Shapes;

                    // Copy each shape from the source worksheet to the destination worksheet
                    foreach (Shape srcShape in srcShapes)
                    {
                        try
                        {
                            // Preserve original position using shape bounds
                            int topRow = srcShape.UpperLeftRow;
                            int leftColumn = srcShape.UpperLeftColumn;
                            int bottomRow = srcShape.LowerRightRow;
                            int rightColumn = srcShape.LowerRightColumn;

                            // AddCopy requires the target cell range for the shape
                            destShapes.AddCopy(srcShape, topRow, leftColumn, bottomRow, rightColumn);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to copy shape '{srcShape.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the workbook with the newly created worksheets that now contain the copied drawing objects
                templateWorkbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}