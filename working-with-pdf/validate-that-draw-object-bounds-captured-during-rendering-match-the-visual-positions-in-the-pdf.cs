// Title: Validate Excel drawing object bounds against PDF rendering positions using Aspose.Cells for .NET
// Common Searches: how to get shape left, top, width, height using Aspose.Cells in C# | compare Excel shape coordinates with the positions in the generated PDF using Aspose.Cells | validate drawing object bounds after saving a worksheet as PDF in .NET | retrieve drawing object dimensions before PDF conversion with Aspose.Cells | check if shape bounds captured in Excel match PDF rendering output C#
// Tags: Aspose.Cells shape bounds extraction | Excel drawing object position validation | C# PDF rendering shape verification | Aspose.Cells worksheet to PDF coordinate check | Aspose.Cells shape dimension retrieval

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, accesses the first worksheet, extracts the Left, Top, Width, and Height of the first drawing shape, prints these bounds, and then renders the worksheet to a PDF stream using Aspose.Cells, handling any errors that may occur during shape retrieval or PDF conversion.
class DrawObjectBoundsValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one drawing object (e.g., a shape or picture)
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No drawing objects found on the worksheet.");
                return;
            }

            // Capture bounds of the first drawing object (points: 1 point = 1/72 inch)
            Shape shape = sheet.Shapes[0];
            try
            {
                float left = shape.Left;
                float top = shape.Top;
                float width = shape.Width;
                float height = shape.Height;

                Console.WriteLine("Captured shape bounds from Excel:");
                Console.WriteLine($"Left: {left}, Top: {top}, Width: {width}, Height: {height}");
            }
            catch (Exception shapeEx)
            {
                Console.WriteLine($"Error retrieving shape bounds: {shapeEx.Message}");
            }

            // Render the worksheet to PDF in memory (optional validation step)
            using (MemoryStream pdfStream = new MemoryStream())
            {
                try
                {
                    // Save the entire workbook (or specific worksheet) as PDF to the memory stream
                    workbook.Save(pdfStream, SaveFormat.Pdf);
                    Console.WriteLine("Worksheet rendered to PDF stream successfully.");
                }
                catch (Exception renderEx)
                {
                    Console.WriteLine($"Error during PDF rendering: {renderEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
