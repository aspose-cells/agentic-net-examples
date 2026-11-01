// Title: Generate a QR code placeholder shape linking to a hosted PNG worksheet image with Aspose.Cells for .NET (C#)
// AI Prompts: Add a rectangular shape to the first worksheet, set its Text property to a PNG URL, center the text horizontally and vertically, apply a thin border, and save the workbook. | Create a fallback QR code shape in an Aspose.Cells workbook by inserting a shape, assigning a URL string as its content, and configuring alignment and border settings using C#. | Implement a workaround for missing QR code generation assemblies by placing the QR data as centered text inside a shape and exporting the workbook as an .xlsx file.
// Common Searches: how to insert a QR code placeholder shape in Excel using Aspose.Cells C# | Aspose.Cells C# add shape with URL text as QR code fallback | save Excel workbook with QR code shape linking to PNG image using Aspose.Cells | C# Aspose.Cells generate QR code shape without external QR library | set shape border thickness in Aspose.Cells workbook C#
// Tags: Aspose.Cells insert QR placeholder shape | C# set shape text to URL in Excel | Aspose.Cells configure shape border weight | Excel QR code shape fallback using Aspose.Cells | Aspose.Cells generate QR code without external library

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds a rectangular shape on the first worksheet, sets its Text to a PNG image URL (acting as a QR code placeholder), centers the text, applies a thin border, and saves the file as WorksheetWithQrCode.xlsx.
    class GenerateQrCode
    {
        static void Main()
        {
            try
            {
                // Data to encode in the QR code (can be any string, e.g., a URL)
                string qrData = "https://example.com/worksheet.png";

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangular shape that will hold the QR code (or placeholder text)
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
                Shape qrShape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // base shape type
                    5,                        // upper left row
                    0,                        // upper left column
                    0,                        // top offset (in pixels)
                    0,                        // left offset (in pixels)
                    200,                      // height (in points)
                    200);                     // width (in points)

                // Since Aspose.Cells QR code generation may require additional assemblies,
                // we place the QR data as text inside the shape as a fallback.
                qrShape.Text = qrData;
                qrShape.TextHorizontalAlignment = TextAlignmentType.Center;
                qrShape.TextVerticalAlignment = TextAlignmentType.Center;

                // Adjust the shape's border appearance
                qrShape.Line.Weight = 0.75; // thin border

                // Save the workbook
                string outputPath = "WorksheetWithQrCode.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
