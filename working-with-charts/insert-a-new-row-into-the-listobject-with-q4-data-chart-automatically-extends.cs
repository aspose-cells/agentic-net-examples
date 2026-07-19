// Title: Add a Row to an Aspose.Cells ListObject and Auto‑Expand the Linked Chart (C#)
// Description: Creates a workbook with a two‑column ListObject (Quarter, Sales), adds a column chart that uses structured references, then inserts a Q4 row via PutCellValue. Because the chart series reference the table, the chart automatically expands to include the new data when saved.
// Keywords: Aspose.Cells ListObject add row | dynamic chart Aspose.Cells | structured reference chart C# | C# insert row table chart update | Aspose.Cells column chart expand | PutCellValue ListObject | Aspose.Cells table chart synchronization
// Common Searches: how to insert a row into a ListObject and update a chart in Aspose.Cells | Aspose.Cells chart expands when table grows | C# add Q4 data to table and extend chart automatically | structured references for dynamic charts Aspose.Cells | Aspose.Cells add row to table without recreating chart
// Developer Intent: Insert a new data row into an Aspose.Cells ListObject so the existing chart automatically reflects the added values.
// Use Cases: Append quarterly sales to a table and keep a column chart in sync without rebuilding it. | Generate financial reports where tables and charts stay linked as rows are programmatically added. | Build dashboards that programmatically add data rows to a ListObject and instantly update associated charts.
// AI Prompts: Show C# code to add a Q4 row to an Aspose.Cells ListObject and have the linked chart auto‑expand. | Explain how structured references let a chart grow with a ListObject in Aspose.Cells. | Provide a step‑by‑step example of inserting a row into a table and updating a column chart using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with a two‑column ListObject (Quarter, Sales), adds a column chart that uses structured references, then inserts a Q4 row via PutCellValue. Because the chart series reference the table, the chart automatically expands to include the new data when saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate source data ----------
                sheet.Cells["A1"].PutValue("Quarter");   // Header for column 0
                sheet.Cells["B1"].PutValue("Sales");     // Header for column 1

                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(120);

                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(150);

                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(180);

                // ---------- Create a ListObject (table) covering the data ----------
                int tableIndex = sheet.ListObjects.Add("A1", "B4", true);   // Add(startCell, endCell, hasHeaders)
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "Table1"; // Set table name for structured references

                // ---------- Add a chart that references the table ----------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);   // Add(ChartType, topRow, leftColumn, bottomRow, rightColumn)
                Chart chart = sheet.Charts[chartIndex];

                // Use structured references so the chart expands when the table grows
                chart.NSeries.Add("=Sheet1!Table1[Sales]", true);               // SeriesCollection.Add(string dataArea, bool isVertical)
                chart.NSeries.CategoryData = "=Sheet1!Table1[Quarter]";

                // ---------- Insert a new row into the ListObject with Q4 data ----------
                int newRowOffset = table.EndRow - table.StartRow + 1;   // Offset for the next row after current data
                table.PutCellValue(newRowOffset, 0, "Q4");             // PutCellValue(rowOffset, columnOffset, value)
                table.PutCellValue(newRowOffset, 1, 200);             // Sales value for Q4

                // ---------- Save the workbook ----------
                string outputPath = "ListObjectChart_Q4.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
