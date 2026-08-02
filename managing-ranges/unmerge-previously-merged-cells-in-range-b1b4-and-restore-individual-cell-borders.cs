// Title: C# – Unmerge B1:B4 and Add Thin Black Borders with Aspose.Cells
// Description: Demonstrates how to split a previously merged range B1:B4 using Aspose.Cells.Range.UnMerge, then apply a thin black border to each individual cell and save the workbook as UnmergedBorders.xlsx.
// Keywords: Aspose.Cells unmerge range C# | add cell borders Aspose.Cells | restore borders after unmerge | C# Excel thin black border | Range.UnMerge example
// Common Searches: how to unmerge cells with Aspose.Cells .NET | apply borders to each cell after unmerge C# | Aspose.Cells set thin black border on range | unmerge B1:B4 and keep borders Aspose
// Developer Intent: Split the merged range B1:B4 and give every cell its own thin black border.
// Use Cases: Revert a temporary merged header to separate rows while keeping a printable border layout. | Generate Excel reports where merged cells must be expanded based on user input without losing visual consistency. | Automate cleanup of merged cells in legacy spreadsheets, ensuring each cell retains a uniform border style.
// AI Prompts: Write C# code that uses Aspose.Cells to unmerge B1:B4 and apply thin black borders to each cell. | Explain the steps to loop through a range after calling Range.UnMerge and set individual border styles in Aspose.Cells. | Show how to customize border color and line style for cells that were previously merged using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to split a previously merged range B1:B4 using Aspose.Cells.Range.UnMerge, then apply a thin black border to each individual cell and save the workbook as UnmergedBorders.xlsx.
class UnmergeAndRestoreBorders
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Merge cells B1:B4 to demonstrate unmerging later
            // B column is index 1, rows 0‑3 (zero‑based)
            worksheet.Cells.Merge(0, 1, 4, 1);

            // Unmerge the previously merged range B1:B4 using Range.UnMerge
            AsposeRange range = worksheet.Cells.CreateRange("B1", "B4");
            range.UnMerge();

            // Restore individual thin black borders for each cell in B1:B4
            for (int row = 0; row < 4; row++)
            {
                Cell cell = worksheet.Cells[row, 1]; // column B = index 1
                Style style = cell.GetStyle();

                // Apply thin borders on all four sides
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                // Set border color (optional, black is default)
                style.Borders[BorderType.TopBorder].Color = Color.Black;
                style.Borders[BorderType.BottomBorder].Color = Color.Black;
                style.Borders[BorderType.LeftBorder].Color = Color.Black;
                style.Borders[BorderType.RightBorder].Color = Color.Black;

                cell.SetStyle(style);
            }

            // Save the modified workbook
            workbook.Save("UnmergedBorders.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
