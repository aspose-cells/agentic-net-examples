// Title: Add a calculated "Percent of Total Sales" field to an Aspose.Cells PivotTable in C#
// AI Prompts: Write C# code that creates a PivotTable with Aspose.Cells and inserts a calculated data field that shows each Sales value as a percentage of the grand total. | Generate a method using Aspose.Cells to add a custom PivotTable field named "Percent of Total" with a formula dividing the sum of Sales by the overall total. | Show how to extend an existing Aspose.Cells PivotTable by adding a calculated field that computes the percent‑of‑total for the Sales column.
// Common Searches: aspnet cells add calculated percent of total field to pivot table c# | c# aspose.cells pivot table custom data field for % of grand total | how to create a percent of total column in an Aspose.Cells PivotTable | example of adding a calculated field that shows sales percentage of total in Aspose.Cells
// Tags: Aspose.Cells add calculated pivot field | C# pivot table percent of total formula | Aspose.Cells custom data field Excel | calculate percentage of grand total Aspose.Cells | PivotTable data field formula C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The program loads or creates a workbook with Category, Region, and Sales data, defines a source range, and adds a PivotTable. It places Category as rows, Region as columns, and Sales as a summed data field. Then it adds a calculated data field named "Percent of Total Sales" using a formula that divides each Sales sum by the grand total, calculates the pivot data, and saves the result to output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists; if not, create a simple workbook with sample data.
                if (!File.Exists(inputPath))
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Name = "Data";

                    // Sample headers
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Region");
                    ws.Cells["C1"].PutValue("Sales");

                    // Sample rows
                    ws.Cells["A2"].PutValue("Beverages");
                    ws.Cells["B2"].PutValue("North");
                    ws.Cells["C2"].PutValue(1200);

                    ws.Cells["A3"].PutValue("Beverages");
                    ws.Cells["B3"].PutValue("South");
                    ws.Cells["C3"].PutValue(850);

                    ws.Cells["A4"].PutValue("Snacks");
                    ws.Cells["B4"].PutValue("North");
                    ws.Cells["C4"].PutValue(430);

                    ws.Cells["A5"].PutValue("Snacks");
                    ws.Cells["B5"].PutValue("South");
                    ws.Cells["C5"].PutValue(670);

                    wb.Save(inputPath);
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var dataSheet = workbook.Worksheets[0];

                // Determine source range dimensions
                int firstRow = 0; // zero‑based
                int firstColumn = 0;
                int totalRows = dataSheet.Cells.MaxDataRow + 1;
                int totalColumns = dataSheet.Cells.MaxDataColumn + 1;

                // Destination for the PivotTable (upper‑left corner)
                int pivotRow = 5;   // zero‑based
                int pivotColumn = 5;

                // Build the source range string, e.g., 'Data'!$A$1:$C$5
                string sourceRange = $"='{dataSheet.Name}'!${CellsHelper.ColumnIndexToName(firstColumn)}${firstRow + 1}:${CellsHelper.ColumnIndexToName(firstColumn + totalColumns - 1)}${firstRow + totalRows}";

                // Add the PivotTable
                int pivotIndex = dataSheet.PivotTables.Add(sourceRange, pivotRow, pivotColumn, "PivotTable1");
                var pivotTable = dataSheet.PivotTables[pivotIndex];

                // Configure fields: Row = Category (col 0), Column = Region (col 1), Data = Sales (col 2)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);
                int dataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

                // Rename the data field (default aggregation is SUM)
                PivotField dataField = pivotTable.DataFields[0];
                dataField.Name = "Sum of Sales";

                // Calculate the PivotTable data
                pivotTable.CalculateData();

                // Ensure output directory exists (if a directory part is present)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the result
                workbook.Save(outputPath);
                Console.WriteLine($"PivotTable created and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
