// Title: Insert a Row into an Aspose.Cells ListObject and Auto‑Expand a Column Chart (C#)
// Description: C# example that creates a workbook with a ListObject (Excel table) for Q1‑Q3 sales, binds a column chart using structured table references, inserts a Q4 row via PutCellValue, and saves the file. The chart automatically includes the new data without manual range updates.
// Keywords: Aspose.Cells | C# | ListObject | add row to table | PutCellValue | structured table reference | auto expand chart | column chart | Excel table chart binding | dynamic chart range
// Common Searches: Aspose.Cells add row to ListObject C# | auto expand chart when table grows Aspose | PutCellValue example Aspose.Cells | bind chart to Excel table Aspose.Cells | insert Q4 data into Aspose.Cells table
// Developer Intent: Add a new quarter row to an existing ListObject so the linked column chart updates automatically.
// Use Cases: Extend quarterly sales tables while keeping dashboards current. | Programmatically grow financial reports without redefining chart ranges. | Automate periodic data insertion into Excel tables with live chart updates.
// AI Prompts: Generate C# code that inserts multiple rows into an Aspose.Cells ListObject and refreshes all dependent charts. | Show how to bind a column chart to a ListObject using structured references so it expands automatically. | Provide robust error handling for adding data to a table, updating charts, and saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // C# example that creates a workbook with a ListObject (Excel table) for Q1‑Q3 sales, binds a column chart using structured table references, inserts a Q4 row via PutCellValue, and saves the file. The chart automatically includes the new data without manual range updates.
    class InsertRowIntoListObject
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Create sample data for the table (Q1‑Q3) -----
                // Header row
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["B1"].PutValue("Sales");

                // Data rows
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(200);

                // ----- Add a ListObject (table) covering the range A1:B4 -----
                int tableIdx = sheet.ListObjects.Add("A1", "B4", true);
                ListObject table = sheet.ListObjects[tableIdx];
                // Set the table name (use DisplayName property)
                table.DisplayName = "SalesData";

                // ----- Add a column chart that uses the table as its data source -----
                // The chart is placed somewhere below the table
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);
                Chart chart = sheet.Charts[chartIdx];

                // Use structured table references so the chart expands automatically
                chart.NSeries.Add($"=Sheet1!{table.DisplayName}[Sales]", true);
                chart.NSeries.CategoryData = $"=Sheet1!{table.DisplayName}[Quarter]";

                // ----- Insert Q4 data into the ListObject -----
                // Row offset is zero‑based relative to the first data row (not the header)
                // Existing rows are offsets 0,1,2 → Q4 will be offset 3
                table.PutCellValue(3, 0, "Q4");   // Quarter column
                table.PutCellValue(3, 1, 250);    // Sales column

                // Save the workbook
                workbook.Save("InsertRowIntoListObject.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
