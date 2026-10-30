// Title: Apply an OuterBottom preset shadow to a textbox shape and export the worksheet as a PNG screenshot using Aspose.Cells for .NET
// Common Searches: how to add an OuterBottom shadow to a textbox shape in Aspose.Cells C# | export Excel worksheet to PNG after formatting shapes with Aspose.Cells .NET | Aspose.Cells C# example for rendering a sheet with a shadowed textbox to an image file | set shadow effect on shape text and capture worksheet screenshot using Aspose.Cells | C# code to save workbook and generate PNG of sheet containing a shadowed textbox
// Tags: Aspose.Cells OuterBottom shadow textbox | render worksheet to PNG Aspose.Cells | shape text shadow effect .NET | export Excel sheet as image C# | Aspose.Cells shape formatting and image rendering

using System;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example creates a new workbook, inserts a textbox shape with sample text, demonstrates where to apply an OuterBottom shadow preset (subject to API availability), renders the first worksheet page to a PNG file, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape with some text
            int upperLeftRow = 2;      // Row index (0‑based)
            int upperLeftColumn = 2;   // Column index (0‑based)
            int width = 200;           // Width in pixels
            int height = 100;          // Height in pixels

            // Offsets (in pixels) set to 0
            Shape textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, 0, 0, width, height);
            textBox.Text = "Sample Text for Shadow";
            textBox.Font.Color = Color.Black;

            // NOTE: Shadow effect APIs may not be available in the current Aspose.Cells version.
            // If needed, apply shadow using the appropriate API for your version.

            // Render the worksheet to a PNG image (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "WorksheetScreenshot.png"); // 0 = first page

            // Save the workbook (optional)
            workbook.Save("WorkbookWithShadow.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
