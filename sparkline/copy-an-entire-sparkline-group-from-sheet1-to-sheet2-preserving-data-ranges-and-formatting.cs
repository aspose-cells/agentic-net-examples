// Title: How to copy a sparkline group from Sheet1 to Sheet2 while preserving data ranges and formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that duplicates an existing sparkline group to another worksheet, updating the data range reference and copying all visual formatting. | Create a .NET snippet that copies cell values, styles, and a sparkline group from one sheet to a new sheet, preserving high‑point, low‑point, series colors, and line weight.
// Common Searches: Aspose.Cells copy sparkline group to another worksheet C# example | preserve sparkline formatting when moving between sheets Aspose.Cells .NET | update sparkline data range after copying to a new sheet using Aspose.Cells | duplicate line sparkline group programmatically in C# | how to clone sparkline group with formatting in Aspose.Cells workbook
// Tags: Aspose.Cells sparkline group transfer | clone sparkline visual settings C# | update sparkline data range worksheet | duplicate sparkline group Aspose.Cells | preserve sparkline style during sheet copy .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineGroupCopyDemo
{
    // The example creates a workbook, adds a line sparkline group on Sheet1, copies the source data and cell styles to Sheet2, recreates the sparkline group on Sheet2 with the same location, adjusts the data range to reference Sheet2, copies all formatting properties (high‑point, low‑point, series colors, line weight), and saves the file as CopySparklineGroupDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet (Sheet1)
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Populate sample data in Sheet1 (A1:D1)
                sheet1.Cells["A1"].PutValue(5);
                sheet1.Cells["B1"].PutValue(2);
                sheet1.Cells["C1"].PutValue(1);
                sheet1.Cells["D1"].PutValue(3);

                // Define the location where the sparkline will be placed (E1)
                CellArea location = new CellArea
                {
                    StartColumn = 4, // column E (0‑based index)
                    EndColumn = 4,
                    StartRow = 0,    // row 1
                    EndRow = 0
                };

                // Add a sparkline group to Sheet1
                int srcGroupIdx = sheet1.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup srcGroup = sheet1.SparklineGroups[srcGroupIdx];

                // Optional: set some formatting on the source group
                srcGroup.ShowHighPoint = true;
                srcGroup.ShowLowPoint = true;
                srcGroup.HighPointColor = workbook.CreateCellsColor();
                srcGroup.HighPointColor.Color = Color.Green;
                srcGroup.LowPointColor = workbook.CreateCellsColor();
                srcGroup.LowPointColor.Color = Color.Red;
                srcGroup.SeriesColor = workbook.CreateCellsColor();
                srcGroup.SeriesColor.Color = Color.Blue;
                srcGroup.LineWeight = 1.0;

                // ------------------------------------------------------------
                // Create Sheet2 and copy the data needed for the sparkline
                // ------------------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet2.Name = "Sheet2";

                // Copy cell values and formats from Sheet1 to Sheet2
                int maxRow = sheet1.Cells.MaxDataRow;
                int maxCol = sheet1.Cells.MaxDataColumn;
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell srcCell = sheet1.Cells[row, col];
                        Cell destCell = sheet2.Cells[row, col];
                        destCell.PutValue(srcCell.Value);
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }

                // ------------------------------------------------------------
                // Prepare parameters for the new sparkline group on Sheet2
                // ------------------------------------------------------------
                SparklineType type = srcGroup.Type;

                // Adjust data range to refer to Sheet2
                string srcDataRange = srcGroup.Sparklines[0].DataRange; // e.g., "Sheet1!A1:D1"
                string dataRange = srcDataRange.Replace(sheet1.Name, sheet2.Name);

                // Orientation – the source group was created with isVertical = false
                bool isVertical = false;

                // Location range – derived from the first and last sparkline positions
                int startRow = srcGroup.Sparklines[0].Row;
                int startCol = srcGroup.Sparklines[0].Column;
                int endRow = srcGroup.Sparklines[srcGroup.Sparklines.Count - 1].Row;
                int endCol = srcGroup.Sparklines[srcGroup.Sparklines.Count - 1].Column;

                CellArea destLocation = new CellArea
                {
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = startCol,
                    EndColumn = endCol
                };

                // ------------------------------------------------------------
                // Add the sparkline group to Sheet2 using the same parameters
                // ------------------------------------------------------------
                int destGroupIdx = sheet2.SparklineGroups.Add(type, dataRange, isVertical, destLocation);
                SparklineGroup destGroup = sheet2.SparklineGroups[destGroupIdx];

                // ------------------------------------------------------------
                // Copy formatting from the source group to the destination group
                // ------------------------------------------------------------
                destGroup.ShowHighPoint = srcGroup.ShowHighPoint;
                destGroup.ShowLowPoint = srcGroup.ShowLowPoint;
                destGroup.HighPointColor = srcGroup.HighPointColor;
                destGroup.LowPointColor = srcGroup.LowPointColor;
                destGroup.SeriesColor = srcGroup.SeriesColor;
                destGroup.LineWeight = srcGroup.LineWeight;

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                workbook.Save("CopySparklineGroupDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
