// Title: C# unit test for verifying rectangle shape text replacement with Aspose.Cells
// Common Searches: how to assert shape text after saving workbook using Aspose.Cells C# | unit testing Aspose.Cells shape Text property persistence | C# example testing rectangle shape text replacement in Excel file | verify Aspose.Cells shape text remains after workbook reload | write MSTest or NUnit test for Aspose.Cells shape text change
// Tags: Aspose.Cells shape text replacement unit test | rectangle shape Text property verification C# | Excel workbook save reload shape persistence | memory stream validation of Aspose.Cells workbook | MsoDrawingType.Rectangle text assignment test

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtDemo
{
    // The example creates a new Workbook, adds a rectangle shape, sets its Text to "Replaced", saves the workbook to a MemoryStream as XLSX, reloads it, retrieves the shape, and asserts that the loaded shape's Text matches the expected value, throwing an exception on mismatch.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ShapeTextReplacement_ShouldReplaceText();
                Console.WriteLine("Shape text replacement test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        static void ShapeTextReplacement_ShouldReplaceText()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add a simple rectangle shape to the worksheet
                // Parameters: drawing type, upper left row, column, row offset, column offset, height, width
                var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 50, 100);

                // Replace the shape's text with a new value
                const string newText = "Replaced";
                shape.Text = newText;

                // Save to a memory stream and reload to verify persistence
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    ms.Position = 0;

                    var loadedWorkbook = new Workbook(ms);
                    var loadedShape = loadedWorkbook.Worksheets[0].Shapes[0];

                    if (loadedShape == null)
                        throw new InvalidOperationException("Loaded shape is null.");

                    // Verify that the shape's text matches the new value
                    if (!string.Equals(loadedShape.Text, newText, StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            $"Shape text mismatch. Expected: '{newText}', Actual: '{loadedShape.Text}'.");
                }
            }
            catch (Exception ex)
            {
                // Rethrow to be caught by the outer handler with a clear message
                throw new Exception($"Shape text replacement failed: {ex.Message}", ex);
            }
        }
    }
}
