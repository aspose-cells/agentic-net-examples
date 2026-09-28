// Title: Copy rows with embedded images and free‑floating shapes using Aspose.Cells Cells.CopyRows in C#
// AI Prompts: Use Aspose.Cells Cells.CopyRows to duplicate rows 0‑2 from a source worksheet to another worksheet while keeping any pictures and rectangle shapes intact. | Transfer a range of rows that contain embedded pictures and drawing objects to a different workbook using the default copy behavior of Aspose.Cells in C#. | Copy rows with free‑floating shapes and images from one worksheet to a new workbook without manually handling the drawing objects, leveraging Aspose.Cells CopyRows method.
// Common Searches: aspnet copy rows with pictures using Aspose.Cells Cells.CopyRows | how to preserve shapes when copying rows between worksheets in C# Aspose.Cells | copy rows including embedded images to another workbook Aspose.Cells .NET | default CopyRows behavior retains drawing objects Aspose.Cells example | C# Aspose.Cells copy rows 0-2 to row 5 with images and shapes
// Tags: copy rows preserving images Aspose.Cells | retain drawing objects when copying rows .NET | default Cells.CopyRows includes shapes | move rows between workbooks C# Aspose.Cells | duplicate rows with embedded pictures Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRowCopyWithImages
{
    // The example creates a source workbook, adds text, an optional image, and a rectangle shape, then copies rows 0‑2 to row 5 of a new workbook using Cells.CopyRows. The default copy behavior automatically preserves the embedded picture and free‑floating shape, and the destination workbook is saved as DestinationWithCopiedRows.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create source workbook and add sample data, image and shape
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Fill some data in rows 0-2
                sourceSheet.Cells["A1"].PutValue("Row 1");
                sourceSheet.Cells["A2"].PutValue("Row 2");
                sourceSheet.Cells["A3"].PutValue("Row 3");

                // Add an image if the file exists
                string imagePath = "sample_image.png";
                if (File.Exists(imagePath))
                {
                    // The picture will be anchored to cell B2 (row index 1, column index 1)
                    int pictureIndex = sourceSheet.Pictures.Add(1, 1, imagePath);
                    Picture picture = sourceSheet.Pictures[pictureIndex];
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture addition.");
                }

                // Add a simple rectangle shape anchored to cell C1
                ShapeCollection shapes = sourceSheet.Shapes;
                RectangleShape rect = shapes.AddRectangle(0, 0, 0, 0, 100, 50);
                rect.Placement = PlacementType.FreeFloating;

                // Create destination workbook (empty)
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Copy rows 0-2 from source to destination starting at row 5
                // This uses the default copy behavior which also copies drawing objects
                destSheet.Cells.CopyRows(sourceSheet.Cells, 0, 5, 3);

                // Save the destination workbook to verify that rows, images and shapes are copied
                destWorkbook.Save("DestinationWithCopiedRows.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
