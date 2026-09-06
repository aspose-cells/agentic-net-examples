// Title: Apply the workbook's Accent1 theme color as a solid background to all used cells in the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with the workbook's Accent1 theme color and applies it to every used cell in the first worksheet via Aspose.Cells. | Show how to retrieve a theme color with GetThemeColor and set it as the foreground color of a solid‑fill style for a cell range in Aspose.Cells. | Provide a snippet that loops through the used range of the first sheet and assigns the Accent1‑based style to each cell.
// Common Searches: Aspose.Cells C# set background of all cells to workbook Accent1 theme color | How to use GetThemeColor to fill cells with theme accent in Aspose.Cells .NET | Apply solid theme accent fill to used range of first worksheet Aspose.Cells | C# code example for styling entire worksheet with workbook theme color Aspose.Cells
// Tags: set cell background using GetThemeColor Aspose.Cells | apply Accent1 theme color to worksheet cells C# | solid fill style with workbook theme color Aspose.Cells | format used range with theme accent Aspose.Cells .NET | create style from workbook theme color Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, builds a style whose foreground color is the workbook's Accent1 theme color with a solid fill, iterates over the used cells of the first worksheet to apply this style, and saves the workbook as Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // {CreateWorkbook}

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the used range of the worksheet
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxColumn = worksheet.Cells.MaxDataColumn;

                // Create a style that uses the custom Accent1 theme color for the background
                Style accentStyle = workbook.CreateStyle();
                // Use GetThemeColor to obtain the Accent1 theme color
                accentStyle.ForegroundColor = workbook.GetThemeColor(ThemeColorType.Accent1);
                accentStyle.Pattern = BackgroundType.Solid;

                // Apply the style to every used cell in the first worksheet
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        cell.SetStyle(accentStyle);
                    }
                }

                // Save the workbook
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath); // {SaveWorkbook}
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
