using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDiffDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a workbook and two worksheets to compare
                // -------------------------------------------------
                Workbook workbook = new Workbook();

                // Worksheet 1 – fill with sample data
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("ID");
                sheet1.Cells["B1"].PutValue("Name");
                sheet1.Cells["A2"].PutValue(1);
                sheet1.Cells["B2"].PutValue("Alice");
                sheet1.Cells["A3"].PutValue(2);
                sheet1.Cells["B3"].PutValue("Bob");

                // Worksheet 2 – fill with slightly different data
                int sheet2Index = workbook.Worksheets.Add();
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Sheet2";
                sheet2.Cells["A1"].PutValue("ID");
                sheet2.Cells["B1"].PutValue("Name");
                sheet2.Cells["A2"].PutValue(1);
                sheet2.Cells["B2"].PutValue("Alice");
                sheet2.Cells["A3"].PutValue(2);
                sheet2.Cells["B3"].PutValue("Robert"); // Different value

                // -------------------------------------------------
                // 2. Enumerate cells of each worksheet and store values in dictionaries
                // -------------------------------------------------
                var valuesSheet1 = new Dictionary<string, string>();
                var valuesSheet2 = new Dictionary<string, string>();

                // Enumerate cells in Sheet1
                IEnumerator enum1 = sheet1.Cells.GetEnumerator();
                while (enum1.MoveNext())
                {
                    Cell cell = (Cell)enum1.Current;
                    valuesSheet1[cell.Name] = cell.Value?.ToString() ?? string.Empty;
                }

                // Enumerate cells in Sheet2
                IEnumerator enum2 = sheet2.Cells.GetEnumerator();
                while (enum2.MoveNext())
                {
                    Cell cell = (Cell)enum2.Current;
                    valuesSheet2[cell.Name] = cell.Value?.ToString() ?? string.Empty;
                }

                // -------------------------------------------------
                // 3. Create a diff report worksheet
                // -------------------------------------------------
                int diffSheetIndex = workbook.Worksheets.Add();
                Worksheet diffSheet = workbook.Worksheets[diffSheetIndex];
                diffSheet.Name = "DiffReport";

                // Header row
                diffSheet.Cells["A1"].PutValue("Cell");
                diffSheet.Cells["B1"].PutValue("Sheet1 Value");
                diffSheet.Cells["C1"].PutValue("Sheet2 Value");

                // Style for mismatched rows (yellow background)
                Style diffStyle = workbook.CreateStyle();
                diffStyle.ForegroundColor = Color.Yellow;
                diffStyle.Pattern = BackgroundType.Solid;
                StyleFlag styleFlag = new StyleFlag { CellShading = true };

                int diffRowIndex = 1; // zero‑based index; start after header

                // -------------------------------------------------
                // 4. Compare values and write mismatches to the report
                // -------------------------------------------------
                var allKeys = valuesSheet1.Keys.Union(valuesSheet2.Keys);
                foreach (string cellName in allKeys)
                {
                    valuesSheet1.TryGetValue(cellName, out string val1);
                    valuesSheet2.TryGetValue(cellName, out string val2);

                    // If values differ (including case where one is missing)
                    if (!string.Equals(val1, val2, StringComparison.Ordinal))
                    {
                        diffSheet.Cells[diffRowIndex, 0].PutValue(cellName);
                        diffSheet.Cells[diffRowIndex, 1].PutValue(val1);
                        diffSheet.Cells[diffRowIndex, 2].PutValue(val2);

                        // Apply highlight style to the entire row
                        diffSheet.Cells.CreateRange(diffRowIndex, 0, 1, 3).ApplyStyle(diffStyle, styleFlag);

                        diffRowIndex++;
                    }
                }

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                string outputPath = "WorksheetDiffReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Diff report saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}