// Title: Apply Diagonal Stripe Fill Pattern to a Cell Range with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, selects the first worksheet, and iterates through a defined range (e.g., B2:D4). For each cell a style is built, the BackgroundType is set to DiagonalStripe, foreground and background colors are assigned, the style is applied, and the file is saved as DiagonalStripeHighlight.xlsx.
// Keywords: Aspose.Cells C# fill pattern | DiagonalStripe background | Excel cell style .NET | highlight cells with pattern | BackgroundType.DiagonalStripe | Excel formatting Aspose
// Common Searches: how to set diagonal stripe background in Aspose.Cells | apply fill pattern to a range using Aspose.Cells C# | Aspose.Cells style diagonal stripes example | C# code for patterned cell background in Excel
// Developer Intent: Add a diagonal‑stripe background to a specific range of cells in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Visually emphasize a block of data in a financial report. | Mark cells that meet a custom condition (e.g., values above a threshold) with a patterned background. | Create a styled header or footer row that stands out without using solid colors. | Design printable worksheets where patterned fills improve readability for color‑blind users.
// AI Prompts: Generate C# code that applies a diagonal stripe fill pattern to any user‑defined cell range with Aspose.Cells. | Show how to reuse a single Style object for bulk cell formatting to improve performance in Aspose.Cells. | Explain how to combine a DiagonalStripe background with custom fonts, borders, and conditional formatting in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsFillPatternDemo
{
    // Creates a new workbook, selects the first worksheet, and iterates through a defined range (e.g., B2:D4). For each cell a style is built, the BackgroundType is set to DiagonalStripe, foreground and background colors are assigned, the style is applied, and the file is saved as DiagonalStripeHighlight.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range of cells to highlight (e.g., B2:D4)
            string startCell = "B2";
            string endCell = "D4";

            // Loop through each cell in the range and apply diagonal stripe pattern
            for (int row = worksheet.Cells[startCell].Row; row <= worksheet.Cells[endCell].Row; row++)
            {
                for (int col = worksheet.Cells[startCell].Column; col <= worksheet.Cells[endCell].Column; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Create a new style for the cell
                    Style style = workbook.CreateStyle();

                    // Set the background pattern to diagonal stripe
                    style.Pattern = BackgroundType.DiagonalStripe;

                    // Set foreground and background colors for the pattern
                    style.ForegroundColor = Color.Yellow;   // Color of the stripes
                    style.BackgroundColor = Color.Blue;     // Color behind the stripes

                    // Apply the style to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the workbook (save rule)
            workbook.Save("DiagonalStripeHighlight.xlsx", SaveFormat.Xlsx);
        }
    }
}
