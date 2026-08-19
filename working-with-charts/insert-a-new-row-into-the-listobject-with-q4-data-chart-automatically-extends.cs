// Title: C# – Insert a Row into an Aspose.Cells ListObject and Auto‑Expand the Linked Column Chart
// Description: Demonstrates how to create a workbook with a ListObject (SalesTable), bind a column chart to the table using structured references, add a new quarter row programmatically, and have the chart automatically include the new data when the file is saved.
// Keywords: Aspose.Cells ListObject add row C# | Aspose.Cells chart auto expand | structured reference chart Aspose.Cells | C# insert data into table Aspose.Cells | dynamic chart range Aspose.Cells
// Common Searches: how to add a row to a ListObject in Aspose.Cells and update the chart | Aspose.Cells C# chart expands when table grows | using structured references for dynamic charts in Aspose.Cells | append quarterly data to Aspose.Cells table and refresh chart
// Developer Intent: Programmatically append a new data row to a ListObject so the existing column chart updates automatically without manual range changes.
// Use Cases: Quarterly sales reporting where each new quarter is added to a table and the chart reflects it instantly. | Financial dashboards that keep tables and associated charts synchronized as data rows are appended. | Automated data‑entry pipelines that maintain up‑to‑date visualizations in Excel workbooks generated with Aspose.Cells.
// AI Prompts: Generate C# code to add multiple rows to an Aspose.Cells ListObject and ensure all linked charts expand accordingly. | Show how to bind a chart to a ListObject with structured references in Aspose.Cells, then insert a new data row programmatically. | Explain how to verify the series range of a chart after extending a ListObject in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // For ListObject
using Aspose.Cells.Charts;   // For Chart and ChartType

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with a ListObject (SalesTable), bind a column chart to the table using structured references, add a new quarter row programmatically, and have the chart automatically include the new data when the file is saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate initial data for the table (Quarter vs Sales)
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(200);

                // Add a ListObject (table) that covers the data range A1:B4
                int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Set a recognizable name for the table (use DisplayName property)
                table.DisplayName = "SalesTable";

                // Add a column chart that uses the table as its data source
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Use structured references so the chart will grow automatically when the table expands
                chart.NSeries.Add("=Sheet1!SalesTable[Sales]", true);
                chart.NSeries.CategoryData = "=Sheet1!SalesTable[Quarter]";

                // Insert a new row into the ListObject with Q4 data
                // Row offset is relative to the table start (0 = header row). Existing data rows are offsets 1‑3.
                // Offset 4 adds the next row after the current data.
                table.PutCellValue(4, 0, "Q4");   // Quarter column
                table.PutCellValue(4, 1, 250);   // Sales column

                // Define output file path
                string outputPath = "QuarterlySales.xlsx";

                // Save the workbook (the chart will automatically include the new Q4 row)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
