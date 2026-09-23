// Title: Create a rectangle shape with a worksheet hyperlink for intra‑workbook navigation using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to Sheet1, sets its display text, and links it to cell A1 of Sheet2 with Aspose.Cells. | Write a snippet that applies a solid line style to a shape, attaches an intra‑workbook hyperlink, and saves the workbook as an .xlsx file. | Provide an example of creating a clickable shape that jumps to another worksheet in the same Excel file using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# add shape hyperlink to another worksheet | how to create a clickable rectangle that navigates to Sheet2 in .xlsx using Aspose.Cells | programmatically link a shape to a cell in a different sheet with Aspose.Cells for .NET | set line style for a shape and add intra‑workbook hyperlink in C# Aspose.Cells | save workbook after adding shape hyperlink with Aspose.Cells API
// Tags: add rectangle shape with worksheet hyperlink Aspose.Cells | intra‑workbook shape navigation C# | shape hyperlink to cell A1 Aspose.Cells | style rectangle shape line Aspose.Cells | save workbook with shape hyperlink .xlsx

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds Sheet1 and Sheet2, inserts a rectangle shape on Sheet1 labeled "Go to Sheet2", attaches a hyperlink to Sheet2!A1, applies a solid line style, and saves the file as HyperlinkShape.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Add a rectangle shape to Sheet1
            Shape shape = sheet1.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1, 0,                     // upper left row, column
                1, 0,                     // top, left (in pixels)
                100, 30);                 // height, width (in pixels)

            // Set the shape's display text
            shape.Text = "Go to Sheet2";

            // Add a worksheet hyperlink to the shape (pointing to Sheet2!A1)
            shape.AddHyperlink("'Sheet2'!A1");

            // Optional styling for the shape
            shape.Line.Weight = 1;
            shape.Line.DashStyle = MsoLineDashStyle.Solid;
            // Fill color can be set if the API supports it; omitted here to avoid compilation issues
            // shape.Fill.SetSolidColor(Color.LightBlue);

            // Save the workbook
            string outputPath = "HyperlinkShape.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
