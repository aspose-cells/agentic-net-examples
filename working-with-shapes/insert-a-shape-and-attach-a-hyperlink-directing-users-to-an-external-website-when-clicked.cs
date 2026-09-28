// Title: Insert a rectangle shape with a clickable hyperlink to an external website using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a rectangle shape at row 2 column 2, set its text to "Visit Aspose", and assign a hyperlink to https://www.aspose.com using Aspose.Cells in C#. | Generate an Excel file that contains a rectangle shape linked to an external URL and save it as an .xlsx file with Aspose.Cells for .NET. | Programmatically place a shape on a worksheet, attach an external hyperlink to the shape, and export the workbook using the Aspose.Cells API in C#.
// Common Searches: how to add a rectangle shape with a hyperlink in an Excel workbook using Aspose.Cells C# | Aspose.Cells example for linking a shape to an external website | C# code to insert a shape and set a clickable URL in Excel with Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | shape hyperlink Aspose.Cells | save workbook with shape hyperlink Xlsx | Aspose.Cells shape insertion example | hyperlinked shape Excel C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a rectangle shape at row 2 column 2, sets its text to "Visit Aspose", assigns a hyperlink to https://www.aspose.com, and saves the file as ShapeWithHyperlink.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a rectangle shape at position (row 2, column 2) with size 100x50 points
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2, // upper left row (zero‑based)
                2, // upper left column (zero‑based)
                0, // top offset in points
                0, // left offset in points
                50, // height in points
                100); // width in points

            // Set the text displayed inside the shape
            shape.Text = "Visit Aspose";

            // Attach a hyperlink that opens an external website when the shape is clicked
            shape.Hyperlink.Address = "https://www.aspose.com";

            // Save the workbook to a file
            string outputPath = "ShapeWithHyperlink.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
