// Title: Unmerge B1:B4 and Add Thin Black Borders with Aspose.Cells for .NET
// Description: Shows how to merge a range, unmerge it, and apply an individual thin black border to each cell in B1:B4 using Aspose.Cells in C#.
// Keywords: Aspose.Cells unmerge range | C# set cell borders | apply thin black border Aspose.Cells | restore individual borders after unmerge | Aspose.Cells .NET border style | Excel unmerge cells programmatically | Aspose.Cells cell style example
// Common Searches: C# Aspose.Cells unmerge cells and add borders | How to apply borders to each cell after unmerge in Aspose.Cells | Aspose.Cells example for unmerging B1:B4 | Set thin black border on a column range using Aspose.Cells | Programmatically restore cell borders after unmerge in .NET
// Developer Intent: The developer needs to split a previously merged range (B1:B4) and give each resulting cell its own thin black border.
// Use Cases: Convert a merged header back to separate rows while keeping a printable grid. | Reset formatting after temporary merged calculations, ensuring each cell displays its own border. | Prepare a worksheet for data entry where merged cells must be split and each cell requires a consistent border style.
// AI Prompts: Generate C# code with Aspose.Cells that unmerges a given range and applies a thin black border to every cell in that range. | Create a reusable method that takes a worksheet, a range address, and a border color, then unmerges the range and sets the specified border on all cells. | Explain how to preserve cell values while unmerging and adding borders using Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUnmergeExample
{
    // Shows how to merge a range, unmerge it, and apply an individual thin black border to each cell in B1:B4 using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge cells B1:B4 for demonstration (rows 0‑3, column 1)
                cells.Merge(0, 1, 4, 1);
                cells[0, 1].PutValue("Merged B1:B4");

                // Unmerge the previously merged range B1:B4
                cells.UnMerge(0, 1, 4, 1);

                // Define a thin black border style
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                // Apply the style to each cell in the range B1:B4
                for (int row = 0; row < 4; row++)
                {
                    cells[row, 1].SetStyle(borderStyle);
                }

                // Save the workbook
                string outputPath = "UnmergedWithBorders.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
