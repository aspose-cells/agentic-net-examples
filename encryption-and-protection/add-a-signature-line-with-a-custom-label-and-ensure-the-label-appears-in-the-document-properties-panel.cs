// Title: Create a signature line with a custom label and add it as a custom document property in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a signature placeholder shape with label in an Excel workbook using Aspose.Cells C# | Aspose.Cells add custom document property for signature label in .NET | C# draw rectangle shape as signature line and set alternative text with Aspose.Cells | save signature label as custom property in Excel file using Aspose.Cells for .NET | Aspose.Cells create signature line and expose label in file properties panel
// Tags: Aspose.Cells insert rectangle shape | shape alternative text Aspose.Cells | Excel custom document property via Aspose.Cells | signature line implementation C# | Aspose.Cells workbook property panel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Excel workbook, adds a rectangular shape positioned on a specific cell to act as a signature line, assigns an alternative text label to the shape, stores the same label as a custom document property named "SignatureLabel", and saves the workbook, making the label visible in the document properties panel.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define position and size for the signature placeholder (row, column, width, height)
            int row = 2;          // Row index (0‑based)
            int column = 1;       // Column index (0‑based)
            int width = 300;      // Width in pixels
            int height = 100;     // Height in pixels

            // Add a rectangular shape to act as a signature line placeholder
            // The shape is positioned at the specified cell with the given size
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                row,                      // Upper‑left row
                column,                   // Upper‑left column
                0,                        // Upper‑left row offset (pixels)
                0,                        // Upper‑left column offset (pixels)
                width,                    // Width (pixels)
                height);                  // Height (pixels)

            // Set a label for the placeholder (appears when the shape is selected)
            shape.AlternativeText = "Approved By: John Doe";

            // Optionally, give the shape a name for easier identification
            shape.Name = "SignaturePlaceholder";

            // Add the same label as a custom document property so it shows in the properties panel
            workbook.CustomDocumentProperties.Add("SignatureLabel", "Approved By: John Doe");

            // Define output file path
            string outputPath = "SignedWorkbook.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
