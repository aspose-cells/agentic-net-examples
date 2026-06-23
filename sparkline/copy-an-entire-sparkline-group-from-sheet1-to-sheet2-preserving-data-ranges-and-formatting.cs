using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineGroupCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook with two worksheets ----------
                Workbook workbook = new Workbook();

                // First worksheet (index 0)
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Add a second sheet where the sparkline group will be copied
                int sheet2Idx = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Idx];
                sheet2.Name = "Sheet2";

                // ---------- Populate sample data on Sheet1 ----------
                // Data for the sparkline (row 1, columns A‑D)
                sheet1.Cells["A1"].PutValue(5);
                sheet1.Cells["B1"].PutValue(2);
                sheet1.Cells["C1"].PutValue(1);
                sheet1.Cells["D1"].PutValue(3);

                // ---------- Create a sparkline group on Sheet1 ----------
                // Location of the sparkline (cell E1)
                CellArea location = CellArea.CreateCellArea("E1", "E1");
                int srcGroupIdx = sheet1.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup srcGroup = sheet1.SparklineGroups[srcGroupIdx];

                // Add a sparkline to the group (the Add on SparklineCollection creates the sparkline item)
                srcGroup.Sparklines.Add(sheet1.Name + "!A1:D1", 0, 4); // row 0, column 4 (E)

                // Set some formatting on the source group (to be copied)
                srcGroup.ShowHighPoint = true;
                srcGroup.ShowLowPoint = true;
                srcGroup.LineWeight = 1.5;
                srcGroup.HighPointColor = workbook.CreateCellsColor();
                srcGroup.HighPointColor.Color = System.Drawing.Color.Green;
                srcGroup.LowPointColor = workbook.CreateCellsColor();
                srcGroup.LowPointColor.Color = System.Drawing.Color.Red;
                srcGroup.SeriesColor = workbook.CreateCellsColor();
                srcGroup.SeriesColor.Color = System.Drawing.Color.Blue;

                // ---------- Copy the entire sparkline group to Sheet2 ----------
                // Create a new group on the destination sheet (type is the same as source)
                int destGroupIdx = sheet2.SparklineGroups.Add(SparklineType.Line);
                SparklineGroup destGroup = sheet2.SparklineGroups[destGroupIdx];

                // Copy each sparkline (data range and location) from source to destination
                foreach (Sparkline sp in srcGroup.Sparklines)
                {
                    // Row and Column are zero‑based indexes; they refer to the cell where the sparkline will appear.
                    destGroup.Sparklines.Add(sp.DataRange, sp.Row, sp.Column);
                }

                // Copy group‑level formatting properties
                destGroup.ShowHighPoint = srcGroup.ShowHighPoint;
                destGroup.ShowLowPoint = srcGroup.ShowLowPoint;
                destGroup.LineWeight = srcGroup.LineWeight;
                destGroup.HighPointColor = srcGroup.HighPointColor;
                destGroup.LowPointColor = srcGroup.LowPointColor;
                destGroup.SeriesColor = srcGroup.SeriesColor;

                // ---------- Save the workbook ----------
                string outputPath = "SparklineGroupCopyResult.xlsx";

                // Ensure the directory exists (in case a relative path is used)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}