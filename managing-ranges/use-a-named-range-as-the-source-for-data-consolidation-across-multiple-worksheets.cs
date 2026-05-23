using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsConsolidationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet 1 – fill sample data and create a name
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "SalesQ1";

                // Header
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Units");

                // Data
                sheet1.Cells["A2"].PutValue("Apple");
                sheet1.Cells["B2"].PutValue(120);
                sheet1.Cells["A3"].PutValue("Banana");
                sheet1.Cells["B3"].PutValue(85);
                sheet1.Cells["A4"].PutValue("Cherry");
                sheet1.Cells["B4"].PutValue(60);

                // Define a named range that covers the data (including header)
                int nameIdx1 = workbook.Worksheets.Names.Add("Q1Data");
                Name q1Name = workbook.Worksheets.Names[nameIdx1];
                q1Name.RefersTo = "=SalesQ1!$A$1:$B$4";

                // -------------------------------------------------
                // Worksheet 2 – fill sample data and create a name
                // -------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets.Add("SalesQ2");

                // Header
                sheet2.Cells["A1"].PutValue("Product");
                sheet2.Cells["B1"].PutValue("Units");

                // Data
                sheet2.Cells["A2"].PutValue("Apple");
                sheet2.Cells["B2"].PutValue(150);
                sheet2.Cells["A3"].PutValue("Banana");
                sheet2.Cells["B3"].PutValue(95);
                sheet2.Cells["A4"].PutValue("Cherry");
                sheet2.Cells["B4"].PutValue(70);

                // Define a named range for the second sheet
                int nameIdx2 = workbook.Worksheets.Names.Add("Q2Data");
                Name q2Name = workbook.Worksheets.Names[nameIdx2];
                q2Name.RefersTo = "=SalesQ2!$A$1:$B$4";

                // -------------------------------------------------
                // Use the named ranges as source for consolidation
                // -------------------------------------------------
                // Build a single source string for consolidation (comma‑separated ranges)
                string consolidationSource = $"{q1Name.RefersTo.TrimStart('=')},{q2Name.RefersTo.TrimStart('=')}";

                // Destination cell for the consolidated pivot table (E5)
                int destRow = 4;      // zero‑based row index
                int destColumn = 4;   // zero‑based column index
                string destCell = CellsHelper.CellIndexToName(destRow, destColumn);

                // Add the pivot table using the consolidation source string
                PivotTableCollection pivots = sheet1.PivotTables;
                int pivotIdx = pivots.Add(consolidationSource, destRow, destColumn, destCell);

                PivotTable pivot = pivots[pivotIdx];
                pivot.Name = "ConsolidatedPivot";

                // Configure the pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");   // Row field – Product
                pivot.AddFieldToArea(PivotFieldType.Data, "Units");    // Data field – Units (sum)

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "ConsolidatedPivotDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}