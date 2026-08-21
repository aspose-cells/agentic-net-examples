// Title: C# – Apply Diagonal Stripe Fill Pattern to Cells with Aspose.Cells
// Description: Shows how to create a style using BackgroundType.DiagonalStripe, set foreground and background colors, apply the style to a cell range (B2:C4), and save the workbook as an XLSX file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | diagonal stripe fill | BackgroundType.DiagonalStripe | cell background pattern | foreground color | background color | style range formatting | Excel export .NET | cell highlighting
// Common Searches: Aspose.Cells set diagonal stripe pattern C# | apply fill pattern to a range with Aspose.Cells | how to change cell background colors Aspose.Cells .NET | create striped cell style Aspose.Cells | save workbook after formatting cells Aspose.Cells
// Developer Intent: The developer wants to highlight a specific range of cells by applying a diagonal‑stripe fill pattern with custom colors using Aspose.Cells for .NET.
// Use Cases: Visually separate sections in a generated report by adding a striped background to header rows. | Draw attention to cells that exceed a threshold, such as budget overruns, with a contrasting stripe pattern. | Mark total or summary rows in financial statements for quick identification.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a diagonal stripe pattern with red stripes on a white background to the range A5:D10. | Create a reusable method that accepts a worksheet, a cell range, foreground and background colors, and applies a diagonal stripe style using Aspose.Cells. | Show an example of applying three different diagonal stripe styles to separate ranges in the same workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using System.Drawing;

// Shows how to create a style using BackgroundType.DiagonalStripe, set foreground and background colors, apply the style to a cell range (B2:C4), and save the workbook as an XLSX file with Aspose.Cells for .NET.
class DiagonalStripeHighlight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional, just to have visible cells)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Create a style with diagonal stripe background pattern
        Style style = workbook.CreateStyle();
        style.Pattern = BackgroundType.DiagonalStripe;   // diagonal stripe pattern
        style.ForegroundColor = Color.Yellow;            // color of the stripes
        style.BackgroundColor = Color.Blue;              // background color

        // Apply the style to the target cells (e.g., B2:C4)
        for (int row = 1; row <= 3; row++)      // rows 2 to 4 (zero‑based index)
        {
            for (int col = 1; col <= 2; col++)  // columns B to C
            {
                sheet.Cells[row, col].SetStyle(style);
            }
        }

        // Save the workbook
        workbook.Save("DiagonalStripeHighlight.xlsx", SaveFormat.Xlsx);
    }
}
