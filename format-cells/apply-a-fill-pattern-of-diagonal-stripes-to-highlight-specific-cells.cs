// Title: How to apply a diagonal stripe fill pattern to a range of cells with Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# program that builds a style with a diagonal stripe fill, sets the foreground to Yellow and the background to LightGray, and applies it to cells A1 through B2 in a new workbook. | Generate an Excel file named DiagonalStripeHighlight.xlsx where the specified range is highlighted with yellow diagonal stripes over a light‑gray base using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# apply pattern fill to a range of cells | how to use BackgroundType enumeration for cell styling in Aspose.Cells | set foreground and background colors in Aspose.Cells style object | C# example highlighting cells with custom fill pattern in Excel | apply style to cells A1 to B2 using Aspose.Cells .NET
// Tags: stripe pattern fill Aspose.Cells C# | set cell style background pattern .NET | cell style color settings Aspose.Cells | apply style to cell range Aspose.Cells | BackgroundType enumeration Aspose.Cells example

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new workbook, defines a style with a diagonal stripe fill (yellow stripes on a light‑gray background), applies the style to cells A1‑B2, and saves the file as DiagonalStripeHighlight.xlsx.
class DiagonalStripeFillExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Example: Fill cells A1 to B2 with diagonal stripe pattern
        // Define the range
        CellArea range = CellArea.CreateCellArea("A1", "B2");

        // Create a style object
        Style stripeStyle = workbook.CreateStyle();

        // Set the fill pattern to diagonal stripes
        stripeStyle.Pattern = BackgroundType.DiagonalStripe;

        // Define foreground (stripe) and background colors
        stripeStyle.ForegroundColor = Color.Yellow;      // Stripe color
        stripeStyle.BackgroundColor = Color.LightGray;   // Base cell color

        // Apply the style to each cell in the range
        for (int row = range.StartRow; row <= range.EndRow; row++)
        {
            for (int col = range.StartColumn; col <= range.EndColumn; col++)
            {
                // Get the cell (creates it if it doesn't exist)
                Cell cell = sheet.Cells[row, col];

                // Apply the style
                cell.SetStyle(stripeStyle);
            }
        }

        // Save the workbook to a file
        workbook.Save("DiagonalStripeHighlight.xlsx");
    }
}
