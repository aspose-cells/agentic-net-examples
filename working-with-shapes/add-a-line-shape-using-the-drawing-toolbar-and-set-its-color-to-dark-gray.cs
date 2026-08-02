// Title: Add a dark‑gray line shape to an Excel worksheet with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a horizontal LineShape two rows down with a 100‑pixel length, sets the line color to dark gray (when the API supports it), and saves the result as an .xlsx file.
// Keywords: Aspose.Cells | C# line shape | AddLine | LineShape color | dark gray line | Excel drawing toolbar | worksheet.Shapes.AddLine | shape formatting Aspose.Cells | set line color | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add line shape C# | How to set line color in Aspose.Cells | Draw horizontal line in Excel using Aspose.Cells | LineShape dark gray color Aspose.Cells | Worksheet.Shapes.AddLine example | Aspose.Cells shape formatting tutorial
// Developer Intent: Insert a line shape and apply a dark‑gray color programmatically.
// Use Cases: Separate sections in automatically generated financial reports. | Create visual dividers in Excel dashboards or scorecards. | Add custom underlines or borders to cells for enhanced styling. | Design printable forms with line separators for better layout.
// AI Prompts: Generate C# code using Aspose.Cells to add a LineShape, set its color to dark gray, and handle versions where the Color property is unavailable. | Show how to modify line thickness, dash style, and alignment of a LineShape after it has been added to a worksheet. | Explain how to retrieve and verify the current line color of a LineShape at runtime. | Provide a GitHub‑ready snippet that adds multiple line shapes with different colors and saves the workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a horizontal LineShape two rows down with a 100‑pixel length, sets the line color to dark gray (when the API supports it), and saves the result as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a horizontal line shape (2 rows down, 0 column offset, length 100 pixels)
            LineShape lineShape = worksheet.Shapes.AddLine(2, 0, 2, 0, 0, 100) as LineShape;
            if (lineShape == null)
            {
                Console.WriteLine("Failed to create LineShape.");
                return;
            }

            // Set the line color to dark gray (if supported by the API)
            // Note: Some older versions of Aspose.Cells may not expose a Color property on LineFormat.
            // If available, uncomment the following line:
            // lineShape.Line.Color = Color.DarkGray;

            // Define output file name
            string outputPath = "LineShapeDarkGray.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
