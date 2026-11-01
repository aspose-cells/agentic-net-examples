// Title: Write a C# unit test to confirm rectangle shape text replacement persists after saving an Aspose.Cells workbook
// AI Prompts: Create an MSTest method that adds a rectangle shape to a worksheet, sets its Text property, saves the workbook to a MemoryStream, reloads it, and asserts that the Text equals the replaced value. | Generate a NUnit test that inserts a rectangle shape, changes its Text, persists the workbook, loads it back, and verifies the new text using Assert.AreEqual.
// Common Searches: how to assert shape text change in Aspose.Cells unit test C# | verify that rectangle shape text is saved in an Aspose.Cells workbook | C# Aspose.Cells unit test for shape Text property persistence | unit testing Aspose.Cells shape text replacement with MSTest | NUnit example for checking shape text after workbook save in Aspose.Cells
// Tags: Aspose.Cells shape text replacement verification | C# unit test workbook persistence Aspose.Cells | MSTest shape Text property assertion | NUnit Aspose.Cells shape validation | memory stream save load Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    // The example creates a workbook, adds a rectangle shape, sets initial text, replaces it, saves the workbook to a MemoryStream, reloads the file, and confirms that the replaced text persists, throwing exceptions if any verification fails.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
                var shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    upperLeftRow: 0,
                    upperLeftColumn: 0,
                    top: 0,
                    left: 0,
                    height: 300,
                    width: 400);

                // Set original text for the shape
                shape.Text = "Original Text";

                // Verify the original text was set correctly
                if (shape.Text != "Original Text")
                    throw new InvalidOperationException("Initial text assignment failed.");

                // Replace the text of the shape
                shape.Text = "Replaced Text";

                // Verify that the text replacement succeeded
                if (shape.Text != "Replaced Text")
                    throw new InvalidOperationException("Shape text replacement did not work as expected.");

                // Save to a memory stream to ensure the workbook can be saved without errors
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0; // Reset stream position for further operations

                    // Load the workbook back to confirm persistence of the changed text
                    var loadedWorkbook = new Workbook(ms);
                    var loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];

                    if (loadedShape.Text != "Replaced Text")
                        throw new InvalidOperationException("Saved shape text does not match expected value.");
                }

                Console.WriteLine("Shape text replacement test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
