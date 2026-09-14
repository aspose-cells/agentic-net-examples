// Title: Add a free‑floating text box at precise cell coordinates with pixel offsets and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, inserts a free‑floating text box at row 4 column 2 with a 15 px X offset and 10 px Y offset, sets the text to blue bold 12 pt, removes the border, and saves the file as a PDF. | Modify the example to place the text box on cell E7, change its size to 250 px width and 50 px height, use red italic 10 pt font, and keep the shape free‑floating when exporting to PDF.
// Common Searches: Aspose.Cells C# how to position a text box using row and column indexes with pixel offsets | place custom text at exact location in PDF generated from Excel with Aspose.Cells | free floating shape placement offset pixels Aspose.Cells .NET example | set text box font color and style in Aspose.Cells before PDF export | remove border from Aspose.Cells shape when saving workbook as PDF
// Tags: add free‑floating graphic Aspose.Cells | pixel offset shape placement C# | export workbook to PDF Aspose.Cells | custom text box font styling Aspose.Cells | remove shape border Aspose.Cells PDF

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPdfTextPlacement
{
    // Demonstrates creating a workbook, adding a free‑floating text box at a specific cell with pixel offsets, customizing its font and border, and saving the result as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the exact position where the custom text will appear.
                // Row and column are zero‑based indexes.
                // Offsets are in pixels relative to the upper‑left corner of the cell.
                int targetRow = 4;          // 5th row (A5)
                int targetColumn = 2;       // 3rd column (C)
                int offsetX = 15;           // 15 pixels to the right inside the cell
                int offsetY = 10;           // 10 pixels down inside the cell

                // Define the size of the text box (width & height in pixels)
                int boxWidth = 200;
                int boxHeight = 40;

                // Add a free‑floating text box shape at the specified coordinates
                // (lifecycle rule: create shape, then set its properties)
                Shape textBox = sheet.Shapes.AddTextBox(
                    targetRow,          // Upper left row index
                    targetColumn,       // Upper left column index
                    offsetX,            // Upper left row offset (pixels)
                    offsetY,            // Upper left column offset (pixels)
                    boxWidth,           // Width (pixels)
                    boxHeight);         // Height (pixels)

                // Ensure the shape is not tied to cell movement/resizing
                textBox.Placement = PlacementType.FreeFloating;

                // Set the custom text
                textBox.Text = "Custom Text at Precise Coordinates";

                // Optional: format the text appearance
                textBox.Font.Color = Color.Blue;
                textBox.Font.Size = 12;
                textBox.Font.IsBold = true;

                // Remove the shape border for a cleaner look
                textBox.Line.Weight = 0;

                // Save the workbook as PDF (lifecycle rule: save)
                workbook.Save("CustomText.pdf", SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
