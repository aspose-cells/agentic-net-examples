// Title: C# – Apply a Horizontal Light‑Gray to Dark‑Blue Gradient Fill on a WordArt Watermark using Aspose.Cells
// Description: Creates a new workbook, inserts a WordArt shape with the text "CONFIDENTIAL", sets a horizontal two‑color gradient (light gray → dark blue), makes it 50 % transparent, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | WordArt watermark | gradient fill | horizontal gradient | light gray | dark blue | transparent shape | Excel workbook | PresetWordArtStyle
// Common Searches: Aspose.Cells C# add gradient WordArt watermark | horizontal light gray to dark blue gradient in Excel shape | make WordArt watermark semi transparent with Aspose.Cells | set two‑color gradient for WordArt using .NET | programmatically create confidential watermark in Excel
// Developer Intent: Generate a WordArt watermark and apply a horizontal light‑gray‑to‑dark‑blue gradient with 50 % transparency in an Excel file.
// Use Cases: Mark confidential reports with a subtle color‑shift watermark. | Brand internal templates using corporate colors in a gradient watermark. | Batch‑process multiple worksheets to add consistent gradient watermarks.
// AI Prompts: Show C# code to change the gradient direction of the WordArt watermark to vertical while keeping the same colors. | Demonstrate how to add a three‑color gradient (light gray, medium blue, dark blue) to a WordArt shape with Aspose.Cells. | Explain how to reuse a predefined gradient fill across several WordArt watermarks on different worksheets.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtWatermark
{
    // Creates a new workbook, inserts a WordArt shape with the text "CONFIDENTIAL", sets a horizontal two‑color gradient (light gray → dark blue), makes it 50 % transparent, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape that will act as the watermark
            // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle1,   // base style (can be changed later)
                "CONFIDENTIAL",                     // watermark text
                5, 0,                               // row and top offset
                5, 0,                               // column and left offset
                200, 600);                          // height and width

            // Ensure the shape uses a gradient fill
            wordArt.Fill.FillType = FillType.Gradient;

            // Obtain the GradientFill object
            GradientFill gradientFill = wordArt.Fill.GradientFill;

            // Apply a two‑color gradient: light gray to dark blue, horizontal direction
            gradientFill.SetTwoColorGradient(
                Color.LightGray,    // first (light) color
                Color.DarkBlue,     // second (dark) color
                GradientStyleType.Horizontal, // gradient style
                1);                 // variant (default)

            // Optionally, make the WordArt semi‑transparent to resemble a watermark
            wordArt.Fill.Transparency = 0.5; // 50% transparent

            // Save the workbook
            workbook.Save("WordArtWatermarkGradient.xlsx");
        }
    }
}
