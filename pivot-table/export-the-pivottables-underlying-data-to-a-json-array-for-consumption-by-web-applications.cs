// Title: Export PivotTable Data to JSON with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, build a PivotTable, switch it to tabular layout, capture the used range, and use Aspose.Cells JsonSaveOptions to generate a clean JSON array that can be consumed by web services or client‑side visualizations.
// Keywords: Aspose.Cells C# | PivotTable to JSON | JsonSaveOptions example | export pivot data as JSON | range to JSON Aspose | tabular form PivotTable | web API JSON payload | Excel pivot export .NET
// Common Searches: Aspose.Cells export PivotTable to JSON C# | How to convert PivotTable results to JSON array | JsonSaveOptions for pivot data in .NET | Get underlying PivotTable data as JSON | C# code to output PivotTable as JSON string
// Developer Intent: Generate a JSON array from the result set of a PivotTable for use in web applications or APIs.
// Use Cases: Feed PivotTable analytics to JavaScript charts or data grids. | Expose spreadsheet calculations through a REST endpoint. | Archive pivot results in a lightweight, language‑agnostic format.
// AI Prompts: Write C# code that extracts the displayed range of a PivotTable and serializes it to JSON with Aspose.Cells. | Show how to configure JsonSaveOptions to include headers, keep empty cells, and export all values as strings for a PivotTable range. | Explain how to determine the used range of a PivotTable sheet when the DataRange property is unavailable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsPivotJsonExport
{
    // Demonstrates how to create a workbook, build a PivotTable, switch it to tabular layout, capture the used range, and use Aspose.Cells JsonSaveOptions to generate a clean JSON array that can be consumed by web services or client‑side visualizations.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and populate source data
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Sample data: Category, SubCategory, Amount
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("SubCategory");
                dataSheet.Cells["C1"].PutValue("Amount");

                dataSheet.Cells["A2"].PutValue("Fruit");
                dataSheet.Cells["B2"].PutValue("Apple");
                dataSheet.Cells["C2"].PutValue(120);

                dataSheet.Cells["A3"].PutValue("Fruit");
                dataSheet.Cells["B3"].PutValue("Banana");
                dataSheet.Cells["C3"].PutValue(80);

                dataSheet.Cells["A4"].PutValue("Vegetable");
                dataSheet.Cells["B4"].PutValue("Carrot");
                dataSheet.Cells["C4"].PutValue(50);

                dataSheet.Cells["A5"].PutValue("Vegetable");
                dataSheet.Cells["B5"].PutValue("Tomato");
                dataSheet.Cells["C5"].PutValue(70);

                // -------------------------------------------------
                // 2. Add a worksheet for the PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Define the source range for the pivot table (A1:C5)
                string sourceRange = "=Data!A1:C5";

                // Add the pivot table at cell A1 of the pivot sheet
                int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields: Category (row), SubCategory (row), Amount (data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Layout the pivot table in tabular form for easier JSON conversion
                pivotTable.ShowInTabularForm();

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 3. Obtain the range that contains the pivot result
                // -------------------------------------------------
                // Since PivotTable.DataRange is not available, use the used range of the pivot sheet.
                int startRow = 0; // Pivot starts at A1
                int startColumn = 0;
                int totalRows = pivotSheet.Cells.MaxDataRow + 1;      // +1 because MaxDataRow is zero‑based
                int totalColumns = pivotSheet.Cells.MaxDataColumn + 1;

                Aspose.Cells.Range pivotRange = pivotSheet.Cells.CreateRange(
                    startRow,
                    startColumn,
                    totalRows,
                    totalColumns);

                // -------------------------------------------------
                // 4. Export the pivot range to JSON
                // -------------------------------------------------
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = true,          // First row contains column names
                    ExportAsString = true,        // Export all values as strings (useful for web)
                    ExportEmptyCells = true,      // Preserve empty cells as null
                    Indent = "  "                 // Pretty‑print with two‑space indentation
                };

                string jsonResult = pivotRange.ToJson(jsonOptions);

                // Output the JSON string
                Console.WriteLine("PivotTable JSON Export:");
                Console.WriteLine(jsonResult);

                // -------------------------------------------------
                // 5. (Optional) Save the workbook for verification
                // -------------------------------------------------
                workbook.Save("PivotExportDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
