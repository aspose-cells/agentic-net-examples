// Title: Create a C# unit test that verifies shape text replacement persists after saving and loading an XLSX workbook with Aspose.Cells
// AI Prompts: Write an MSTest method that adds a rectangle shape to a worksheet, sets its Text property, saves the workbook to a MemoryStream in XLSX format, reloads it, and asserts the Text equals the replaced value using Aspose.Cells for .NET. | Generate a NUnit test case that creates a shape, changes its Text, persists the workbook, reloads from a stream, and validates the updated text with Aspose.Cells.
// Common Searches: how to unit test shape text change in Aspose.Cells C# | assert that rectangle shape text is saved in XLSX using Aspose.Cells | C# test for persisting smartart text after workbook reload | Aspose.Cells verify shape text after saving to memory stream | unit testing Aspose.Cells shape text property persistence
// Tags: Aspose.Cells shape text persistence test | C# unit test workbook save load | MSTest shape text verification Aspose.Cells | NUnit Aspose.Cells shape text assertion | memory stream XLSX shape validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    // Demonstrates how to write a C# unit test that adds a rectangle shape, replaces its Text, saves the workbook to a MemoryStream in XLSX format, reloads the workbook, and asserts that the new text persists using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a simple AutoShape (rectangle) as a placeholder for SmartArt
                // Parameters: shape type, upper left row, upper left column,
                // upper left row offset (pixels), upper left column offset (pixels), height (pixels), width (pixels)
                Shape shape = worksheet.Shapes.AddShape(
                    Aspose.Cells.Drawing.MsoDrawingType.Rectangle, // use MsoDrawingType for compatibility
                    0,          // upper left row
                    0,          // upper left column
                    0,          // upper left row offset in pixels
                    0,          // upper left column offset in pixels
                    300,        // height in pixels
                    400);       // width in pixels

                // Set initial text and then replace it
                shape.Text = "Original Text";
                shape.Text = "Replaced Text";

                // Save to a memory stream (XLSX format)
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0; // Reset stream position for reading

                    // Load the workbook from the memory stream
                    Workbook loadedWorkbook = new Workbook(ms);
                    Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

                    // Retrieve the shape from the loaded worksheet
                    Shape loadedShape = loadedWorksheet.Shapes[0] as Shape;
                    if (loadedShape == null)
                    {
                        Console.WriteLine("Shape was not found after loading the workbook.");
                        return;
                    }

                    // Verify that the text replacement persisted
                    if (loadedShape.Text == "Replaced Text")
                    {
                        Console.WriteLine("Shape text replacement persisted successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Verification failed. Expected 'Replaced Text' but got '{loadedShape.Text}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
