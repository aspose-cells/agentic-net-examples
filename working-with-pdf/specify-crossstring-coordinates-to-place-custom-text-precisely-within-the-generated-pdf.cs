// Title: Position a TextBox at exact pixel coordinates in a PDF with Aspose.Cells for .NET
// Description: Creates a workbook, inserts a free‑floating TextBox shape using row/column indices and top‑/left pixel offsets, applies optional styling, and saves the sheet as a PDF so the text appears at the defined location.
// Keywords: Aspose.Cells PDF export | C# textbox pixel offset | free floating shape Aspose.Cells | exact text placement PDF | custom text coordinates .NET | Aspose.Cells shape positioning | PDFSaveOptions Aspose.Cells
// Common Searches: Aspose.Cells place textbox at specific pixel location | C# set exact coordinates for shape before PDF conversion | how to use PlacementType.FreeFloating in Aspose.Cells | pixel‑based positioning of text in PDF generated from Excel | Aspose.Cells precise text layout in PDF
// Developer Intent: Add a free‑floating TextBox with defined pixel offsets and export the workbook so the text is rendered at that exact spot in the PDF.
// Use Cases: Design invoices where the company header must align to a fixed spot on the PDF page. | Generate certificates with a signature line positioned consistently across all pages. | Overlay a disclaimer or watermark at a predetermined location in reports derived from Excel.
// AI Prompts: Write C# code that adds a free‑floating TextBox at top offset 120 and left offset 250 pixels, then saves the workbook as a PDF using Aspose.Cells. | Explain the effect of PlacementType.FreeFloating on shape coordinates during PDF export in Aspose.Cells. | Show how to convert Excel cell dimensions to pixel offsets for accurate shape placement in a PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsCustomTextPosition
{
    // Creates a workbook, inserts a free‑floating TextBox shape using row/column indices and top‑/left pixel offsets, applies optional styling, and saves the sheet as a PDF so the text appears at the defined location.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Add a TextBox shape that will hold the custom text.
            // Parameters: upper left row, upper left column, top offset (pixels),
            // left offset (pixels), height (pixels), width (pixels)
            // ------------------------------------------------------------
            int upperLeftRow = 0;
            int upperLeftColumn = 0;
            int topOffset = 150;    // vertical position from the top of the sheet (pixels)
            int leftOffset = 200;   // horizontal position from the left of the sheet (pixels)
            int height = 50;        // height of the textbox (pixels)
            int width = 300;        // width of the textbox (pixels)

            Shape textBox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn,
                                                   topOffset, leftOffset, height, width);

            // Set the text that will appear in the PDF
            textBox.Text = "Precise positioned custom text";

            // Make the textbox free‑floating so it does not move with cells
            textBox.Placement = PlacementType.FreeFloating;

            // Optional: style the text (font, color, size)
            textBox.Font.IsBold = true;
            textBox.Font.Size = 14;
            textBox.Font.Color = System.Drawing.Color.DarkBlue;

            // ------------------------------------------------------------
            // Configure PDF save options.
            // TextCrossType is kept at its default value; it does not affect positioning.
            // ------------------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF; the textbox will be rendered at the exact
            // coordinates specified above.
            workbook.Save("CustomTextPosition.pdf", pdfOptions);

            Console.WriteLine("PDF generated with custom text positioned precisely.");
        }
    }
}
