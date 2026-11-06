// Title: Change cells with white font to a light gray solid fill using Aspose.Cells for .NET
// Common Searches: how to replace white font cells with light gray background in Excel using Aspose.Cells C# | Aspose.Cells iterate all worksheets and set solid background for specific cells | C# code to apply solid fill pattern to cells based on foreground color with Aspose.Cells | update cell style conditionally and save workbook using Aspose.Cells .NET | set background type solid and color light gray for cells in Aspose.Cells
// Tags: Aspose.Cells conditional cell background C# | set solid fill pattern Aspose.Cells | iterate worksheets modify cell styles .NET | apply light gray background based on font color Aspose.Cells | save workbook after style changes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook, loops through every worksheet and populated cell, checks for cells whose font color is white, applies a solid light gray background fill, and then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            Style style = cell.GetStyle();

                            // Example condition: change background of cells with white foreground
                            if (style.ForegroundColor.ToArgb() == Color.White.ToArgb())
                            {
                                // Apply a solid background pattern with a light gray color
                                style.Pattern = BackgroundType.Solid;
                                style.ForegroundColor = Color.LightGray;

                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
