using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;
using System.Drawing;

namespace AsposeCellsExternalLinkRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from an existing XLSX file
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range of the worksheet
            AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

            // List to store cells that contain external links
            List<Cell> externalLinkCells = new List<Cell>();

            // Iterate through each cell in the used range
            for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
            {
                for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Get precedents of the formula
                        ReferredAreaCollection precedents = cell.GetPrecedents();
                        foreach (ReferredArea area in precedents)
                        {
                            // Check if any precedent is an external link
                            if (area.IsExternalLink)
                            {
                                externalLinkCells.Add(cell);
                                break; // No need to check other precedents for this cell
                            }
                        }
                    }
                }
            }

            // If external link cells were found, create a range that encompasses them
            if (externalLinkCells.Count > 0)
            {
                int minRow = int.MaxValue, minCol = int.MaxValue;
                int maxRow = int.MinValue, maxCol = int.MinValue;

                foreach (Cell c in externalLinkCells)
                {
                    if (c.Row < minRow) minRow = c.Row;
                    if (c.Column < minCol) minCol = c.Column;
                    if (c.Row > maxRow) maxRow = c.Row;
                    if (c.Column > maxCol) maxCol = c.Column;
                }

                // Create the range covering all external link cells
                int totalRows = maxRow - minRow + 1;
                int totalCols = maxCol - minCol + 1;
                AsposeRange externalLinkRange = sheet.Cells.CreateRange(minRow, minCol, totalRows, totalCols);

                // Example manipulation: set a background color to highlight the range
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Yellow;
                style.Pattern = BackgroundType.Solid;
                StyleFlag flag = new StyleFlag();
                flag.CellShading = true;
                externalLinkRange.ApplyStyle(style, flag);

                Console.WriteLine($"External link range created: {externalLinkRange.RefersTo}");
            }
            else
            {
                Console.WriteLine("No external links found in the workbook.");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}