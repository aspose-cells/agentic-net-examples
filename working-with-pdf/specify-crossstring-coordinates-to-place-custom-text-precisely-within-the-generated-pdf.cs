// Title: Add a TextBox at exact column/row offsets and export to PDF with Aspose.Cells for .NET
// Description: Creates a workbook, inserts a free‑floating TextBox shape using column, row, X‑offset, Y‑offset, width and height parameters, sets its text and font, configures PdfSaveOptions (including TextCrossType), and saves the file as a PDF where the box appears at the defined coordinates.
// Keywords: Aspose.Cells textbox placement | exact coordinates PDF export | AddTextBox column row offset | free floating shape Aspose.Cells | PdfSaveOptions TextCrossType
// Common Searches: Aspose.Cells place textbox at specific pixel coordinates | C# add shape with column row offset before PDF conversion | prevent shape resizing with cells Aspose.Cells | control text overflow in PDF using TextCrossType
// Developer Intent: Insert a textbox containing custom text at a precise location in an Excel sheet and generate a PDF that preserves that exact placement.
// Use Cases: Fixed header or label on every page of a generated PDF report | Disclaimer or notice positioned independently of cell layout | Form field captions aligned precisely in a PDF created from an Excel template
// AI Prompts: Write C# code with Aspose.Cells to add a textbox at column 5, row 10, 15‑pixel X offset, 10‑pixel Y offset, then save as PDF. | Explain how PdfSaveOptions.TextCrossType.CrossKeep affects text that exceeds a textbox width during PDF conversion. | Show how to set a shape’s Placement to FreeFloating so it does not move or resize with cells in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsCustomTextPlacement
{
    // Creates a workbook, inserts a free‑floating TextBox shape using column, row, X‑offset, Y‑offset, width and height parameters, sets its text and font, configures PdfSaveOptions (including TextCrossType), and saves the file as a PDF where the box appears at the defined coordinates.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (optional, just to have a sheet)
            sheet.Cells["A1"].PutValue("Demo Sheet");

            // Add a TextBox shape at precise coordinates
            // Parameters: upper left column, upper left row, offsetX, offsetY, width, height (all in pixels)
            // Here we place the text box starting at column 2 (C), row 3 (3), with offsets and size
            TextBox textBox = sheet.Shapes.AddTextBox(2, 2, 30, 20, 250, 60);

            // Set the custom text
            textBox.Text = "Precise positioned custom text";

            // Optional: adjust text formatting
            textBox.Font.Name = "Arial";
            textBox.Font.Size = 12;
            textBox.Font.IsBold = true;
            textBox.Font.Color = System.Drawing.Color.DarkBlue;

            // Ensure the text box does not move/resize with cells
            textBox.Placement = PlacementType.FreeFloating;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Control how text that exceeds the cell width is displayed (optional)
            pdfOptions.TextCrossType = TextCrossType.CrossKeep;

            // Save the workbook as PDF; the TextBox will appear at the specified coordinates
            workbook.Save("CustomTextPlacement.pdf", pdfOptions);

            Console.WriteLine("PDF generated with custom text at precise coordinates.");
        }
    }
}
