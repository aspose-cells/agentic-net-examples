// Title: Add a Rectangle Shape with Custom Text and an External Hyperlink using Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape's caption to "Visit Aspose Documentation", formats the text in blue and underlined, attaches a hyperlink to https://docs.aspose.com/cells/net/, and saves the file as ShapeWithHyperlink.xlsx.
// Keywords: Aspose.Cells add shape C# | Excel rectangle shape hyperlink | Aspose.Cells set shape text color | C# Aspose.Cells external link | Insert shape with hyperlink Aspose | Aspose.Cells shape formatting
// Common Searches: how to add a clickable shape in Excel with Aspose.Cells | Aspose.Cells C# rectangle shape with URL | format shape text and add hyperlink Aspose.Cells | Aspose.Cells add shape and link to website
// Developer Intent: Generate a rectangle shape, apply custom text styling, and bind an external URL to the shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a “Help” button in automated reports that opens the online API guide. | Embed quick‑access links to dashboards or external tools directly within generated workbooks. | Design interactive call‑to‑action shapes in templates that redirect users to web forms or documentation.
// AI Prompts: Write C# code with Aspose.Cells to insert a rectangle shape, set its caption, apply blue underlined formatting, and add a hyperlink to an external URL. | Explain how to calculate shape position and size based on cell coordinates in Aspose.Cells. | Show how to add multiple shapes, each with a different hyperlink, to the same worksheet using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, inserts a rectangle shape on the first worksheet, sets the shape's caption to "Visit Aspose Documentation", formats the text in blue and underlined, attaches a hyperlink to https://docs.aspose.com/cells/net/, and saves the file as ShapeWithHyperlink.xlsx.
    class InsertShapeWithHyperlink
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column,
                // upper left row offset (pixels), upper left column offset (pixels),
                // height (pixels), width (pixels)
                RectangleShape shape = (RectangleShape)sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

                // Set custom text for the shape
                shape.Text = "Visit Aspose Documentation";

                // Add a hyperlink to the shape that points to an external website
                shape.AddHyperlink("https://docs.aspose.com/cells/net/");

                // Format the text: make it blue and underlined
                shape.Font.Color = Color.Blue;
                shape.Font.Underline = FontUnderlineType.Single;

                // Save the workbook to a file
                string outputPath = "ShapeWithHyperlink.xlsx";
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
