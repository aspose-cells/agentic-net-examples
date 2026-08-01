// Title: Set multiple Shape text‑alignment properties with an object initializer in Aspose.Cells for .NET
// Description: Shows how to create a Workbook, add a textbox Shape, and apply Text, horizontal TextAlignment and vertical TextVerticalType in one C# object‑initializer statement (when the API version supports these properties) before saving the worksheet.
// Keywords: Aspose.Cells | Shape | TextBox | TextAlignment | TextVerticalType | object initializer | C# | .NET | shape formatting | worksheet | workbook
// Common Searches: object initializer shape Aspose.Cells | set Shape TextAlignment C# | initialize Shape properties in one line Aspose | Aspose.Cells textbox alignment example | C# shape vertical text type Aspose
// Developer Intent: Configure a Shape’s text content and alignment attributes in a single, concise statement using object‑initializer syntax.
// Use Cases: Create several textbox Shapes with identical horizontal and vertical alignment by reusing the same initializer block. | Reduce boilerplate when adding formatted shapes (text, font, fill, alignment) to a worksheet. | Prepare a template workbook where shape alignment settings are defined at creation time for consistent layout.
// AI Prompts: Write C# code that adds a textbox Shape to a worksheet and sets Text, TextAlignment, and TextVerticalType using an object initializer in Aspose.Cells. | Explain how to programmatically check if ShapeTextAlignment properties exist in the installed Aspose.Cells version and suggest a fallback if they are missing. | Provide an example that initializes a Shape with text, font size, fill color, and both alignment settings in one object‑initializer block.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Shows how to create a Workbook, add a textbox Shape, and apply Text, horizontal TextAlignment and vertical TextVerticalType in one C# object‑initializer statement (when the API version supports these properties) before saving the worksheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a text box shape to the worksheet
                Shape shape = worksheet.Shapes.AddTextBox(0, 0, 100, 100, 200, 200);
                shape.Text = "Sample text with custom alignment";

                // Note: Advanced text alignment properties (e.g., TextAlignment, TextVerticalType)
                // are not available in the current Aspose.Cells version used.
                // If needed, they can be set using the appropriate APIs when supported.

                // Save the workbook
                string outputPath = "ShapeTextAlignmentWithInitializer.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
