// Title: Aspose.Cells for .NET – Build a PivotTable from a ListObject (named table) with automatic range updates
// Description: Demonstrates how to create a workbook, convert a data range into a ListObject named "SalesTable", add a PivotTable on a separate sheet that uses this named table as its source, configure row/column/data fields, refresh the pivot cache, and save the file. The PivotTable automatically expands when rows are added to the table.
// Keywords: Aspose.Cells | C# | PivotTable | ListObject | named table | dynamic range | refresh pivot cache | Excel automation | programmatic pivot table | GitHub example | Aspose.Cells for .NET
// Common Searches: Aspose.Cells create PivotTable from ListObject | C# PivotTable using named table Aspose.Cells | dynamic range PivotTable Aspose.Cells .NET | refresh pivot cache programmatically Aspose.Cells | GitHub Aspose.Cells PivotTable example
// Developer Intent: Generate a PivotTable that references a ListObject so the source range expands automatically with new data.
// Use Cases: Produce a sales summary that updates instantly when new rows are added to the SalesTable. | Design a reusable reporting workbook where multiple PivotTables share the same named table for consistent data refresh. | Create dashboards that require minimal maintenance because the pivot cache is tied to a dynamic ListObject.
// AI Prompts: Write C# code with Aspose.Cells to add a ListObject called 'SalesTable' and build a PivotTable that references it, including field setup and cache refresh. | Explain how to programmatically append rows to a ListObject and automatically refresh all linked PivotTables in Aspose.Cells. | Show how to change the aggregation function (e.g., Sum, Average) of a data field in a PivotTable created from a named table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsPivotTableWithNamedTable
{
    // Demonstrates how to create a workbook, convert a data range into a ListObject named "SalesTable", add a PivotTable on a separate sheet that uses this named table as its source, configure row/column/data fields, refresh the pivot cache, and save the file. The PivotTable automatically expands when rows are added to the table.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data on the first worksheet
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Add header row
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                // Add some sample rows
                dataSheet.Cells["A2"].PutValue("Electronics");
                dataSheet.Cells["B2"].PutValue("Laptop");
                dataSheet.Cells["C2"].PutValue(1200);

                dataSheet.Cells["A3"].PutValue("Electronics");
                dataSheet.Cells["B3"].PutValue("Phone");
                dataSheet.Cells["C3"].PutValue(800);

                dataSheet.Cells["A4"].PutValue("Furniture");
                dataSheet.Cells["B4"].PutValue("Chair");
                dataSheet.Cells["C4"].PutValue(150);

                dataSheet.Cells["A5"].PutValue("Furniture");
                dataSheet.Cells["B5"].PutValue("Table");
                dataSheet.Cells["C5"].PutValue(300);

                // -------------------------------------------------
                // 2. Convert the range into a named table (ListObject)
                // -------------------------------------------------
                // Determine the used range
                Aspose.Cells.Range usedRange = dataSheet.Cells.MaxDisplayRange;

                // Add a ListObject (named table) covering the used range.
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = dataSheet.ListObjects.Add(
                    usedRange.FirstRow,
                    usedRange.FirstColumn,
                    usedRange.RowCount,
                    usedRange.ColumnCount,
                    true);

                // Retrieve the created table and assign a display name
                ListObject table = dataSheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable";

                // -------------------------------------------------
                // 3. Create a worksheet to host the PivotTable
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // -------------------------------------------------
                // 4. Add a PivotTable that uses the named table as its source
                // -------------------------------------------------
                string sourceData = "SalesTable";   // table name
                string destCell = "A1";             // upper‑left corner of the PivotTable
                string pivotName = "SalesPivot";

                // Add the PivotTable
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, destCell, pivotName);
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // -------------------------------------------------
                // 5. Configure the PivotTable fields
                // -------------------------------------------------
                // Row field: Category
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                // Column field: Product
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                // Data field: Sales (sum)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                // Set the aggregation function for the data field (optional, default is Sum)
                pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

                // -------------------------------------------------
                // 6. Refresh and calculate the PivotTable data
                // -------------------------------------------------
                // Refresh the pivot cache and calculate the results
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 7. Save the workbook
                // -------------------------------------------------
                string outputPath = "PivotTableWithNamedTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
