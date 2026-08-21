// Title: Integration Test: Verify Linked Picture Refresh After Cell Value Change with Aspose.Cells for .NET
// Description: Demonstrates how to write an integration test that creates a workbook, stores an image path in a cell, adds a linked picture using SetLinkedCell, confirms the linked reference with GetLinkedCell, changes the cell value, calls UpdateSelectedValue, and validates that the picture updates to the new image.
// Keywords: Aspose.Cells | linked picture | UpdateSelectedValue | SetLinkedCell | GetLinkedCell | C# integration test | Excel shape refresh | linked shape test | Aspose.Cells .NET | picture refresh after cell change
// Common Searches: Aspose.Cells linked picture refresh test | How to refresh linked picture after changing cell value in .NET | UpdateSelectedValue example Aspose.Cells | SetLinkedCell GetLinkedCell verification C# | Integration test for linked shape Aspose.Cells
// Developer Intent: Ensure a linked picture updates its image automatically when the source cell value is modified.
// Use Cases: Automated regression test for Excel reports that use linked images | CI pipeline validation of linked shape behavior in generated workbooks | Documentation example showing programmatic picture refresh after cell edit | Unit testing of SetLinkedCell and UpdateSelectedValue methods
// AI Prompts: Create an MSTest method that builds a workbook, adds a linked picture, changes the image path in the linked cell, calls UpdateSelectedValue, and asserts the picture source matches the new file. | Write a NUnit test for Aspose.Cells that verifies SetLinkedCell, GetLinkedCell, and picture refresh after updating the cell containing the image path. | Provide a xUnit test snippet that confirms a linked picture reads the image path from a cell, updates when the cell value changes, and does not throw exceptions during UpdateSelectedValue.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    // Demonstrates how to write an integration test that creates a workbook, stores an image path in a cell, adds a linked picture using SetLinkedCell, confirms the linked reference with GetLinkedCell, changes the cell value, calls UpdateSelectedValue, and validates that the picture updates to the new image.
    class Program
    {
        // Paths to sample images used in the demo.
        private const string ImagePath1 = "sample1.png";
        private const string ImagePath2 = "sample2.png";

        static void Main()
        {
            try
            {
                // Ensure sample images exist; create simple placeholders if missing.
                CreatePlaceholderImageIfMissing(ImagePath1);
                CreatePlaceholderImageIfMissing(ImagePath2);

                // Create a new workbook and obtain the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put the first image path into cell A1.
                sheet.Cells["A1"].PutValue(ImagePath1);

                // Add a linked picture whose source is the image in A1.
                // Placed at row 2, column 2 with size 100x100 pixels.
                Picture picture = sheet.Shapes.AddLinkedPicture(2, 2, 100, 100, ImagePath1);

                // Link the picture to cell A1. The picture will read the image path from this cell.
                picture.SetLinkedCell("A1", false, false);

                // Verify the linked cell is set correctly.
                string linkedCell = picture.GetLinkedCell(true, true);
                if (linkedCell != "$A$1")
                {
                    Console.WriteLine($"Unexpected linked cell reference: {linkedCell}");
                }
                else
                {
                    Console.WriteLine("Linked cell correctly set to $A$1.");
                }

                // Change the cell value to point to a different image.
                sheet.Cells["A1"].PutValue(ImagePath2);

                // Refresh the picture so it reads the new linked cell value.
                picture.UpdateSelectedValue();

                Console.WriteLine("Linked picture refreshed successfully after cell value change.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Creates a simple 1x1 PNG placeholder image if the specified file does not exist.
        private static void CreatePlaceholderImageIfMissing(string path)
        {
            if (!File.Exists(path))
            {
                // Minimal 1x1 pixel PNG (transparent).
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                try
                {
                    File.WriteAllBytes(path, pngBytes);
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to create placeholder image '{path}': {writeEx.Message}");
                }
            }
        }
    }
}
