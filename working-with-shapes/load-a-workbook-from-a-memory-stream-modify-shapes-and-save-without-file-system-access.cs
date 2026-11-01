// Title: Load an Excel workbook from a byte array, add or modify shapes, and save back to a byte array using Aspose.Cells for .NET
// Common Searches: how to load an Excel file from a byte array with Aspose.Cells and edit worksheet shapes in C# | Aspose.Cells C# add rectangle shape to a worksheet without writing to disk | modify existing shape position and text in a workbook loaded from a memory stream using Aspose.Cells | save modified Excel workbook to a byte array using Aspose.Cells .NET API | process Excel shapes in memory with Aspose.Cells for serverless or cloud functions
// Tags: load workbook from memory stream Aspose.Cells | add rectangle shape worksheet Aspose.Cells | update shape position size text Aspose.Cells | save workbook to byte array Aspose.Cells | shape manipulation without file system Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates loading an Excel workbook from a byte array via a MemoryStream, adding a new rectangle shape, optionally updating the first existing shape's position, size, and text, and then saving the modified workbook back to a byte array—all using Aspose.Cells for .NET without any file‑system access.
public class WorkbookShapeProcessor
{
    /// <param name="inputBytes">The original workbook bytes.</param>
    /// <returns>Byte array of the modified workbook, or null if an error occurs.</returns>
    public byte[] ProcessWorkbook(byte[] inputBytes)
    {
        if (inputBytes == null || inputBytes.Length == 0)
            throw new ArgumentException("Input workbook bytes are null or empty.", nameof(inputBytes));

        try
        {
            // Load workbook from the input byte array using a MemoryStream (no file system access)
            using (var inputStream = new MemoryStream(inputBytes))
            {
                var workbook = new Workbook(inputStream);
                var worksheet = workbook.Worksheets[0];

                // ------------------------------
                // Shape modifications start here
                // ------------------------------

                // Example 1: Add a new rectangle shape
                Shape rectangle = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    5,   // upper left row
                    0,   // upper left column
                    5,   // top offset (points)
                    0,   // left offset (points)
                    200, // width (points)
                    100  // height (points)
                );

                rectangle.Text = "Sample Rectangle";
                rectangle.Placement = PlacementType.FreeFloating;
                // Note: FillColor, LineWeight, and LineColor properties are not available in the current Aspose.Cells version.
                // If needed, they can be set via the Shape's Fill and Line objects in newer versions.

                // Example 2: Modify an existing shape if any exist
                if (worksheet.Shapes.Count > 0)
                {
                    Shape existingShape = worksheet.Shapes[0];

                    // Change its position and size
                    existingShape.Left = 50;   // points from the left edge
                    existingShape.Top = 50;    // points from the top edge
                    existingShape.Width = 150;
                    existingShape.Height = 80;

                    // Update text (all shapes have a Text property)
                    existingShape.Text = "Updated Shape Text";

                    // Change fill color if supported; otherwise this line is omitted.
                    // existingShape.FillColor = Color.LightGreen;
                }

                // ------------------------------
                // Shape modifications end here
                // ------------------------------

                // Save the modified workbook to a new memory stream
                using (var outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context
            throw new InvalidOperationException("Failed to process workbook shapes.", ex);
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            const string inputPath = "sample.xlsx";
            byte[] inputBytes;

            // Load existing file if present; otherwise create a new empty workbook
            if (File.Exists(inputPath))
            {
                inputBytes = File.ReadAllBytes(inputPath);
            }
            else
            {
                var wb = new Workbook();
                using (var ms = new MemoryStream())
                {
                    wb.Save(ms, SaveFormat.Xlsx);
                    inputBytes = ms.ToArray();
                }
            }

            var processor = new WorkbookShapeProcessor();
            byte[] resultBytes = processor.ProcessWorkbook(inputBytes);

            // Write the modified workbook to disk
            const string outputPath = "result.xlsx";
            File.WriteAllBytes(outputPath, resultBytes);
            Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
