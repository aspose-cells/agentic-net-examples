// Title: Map PivotTable drill‑down rows into an existing ListObject (DetailTable) using Aspose.Cells for .NET
// Description: Demonstrates creating a workbook, adding a pivot table, defining a target sheet with an empty ListObject named DetailTable, and using PivotTable.ShowDetail to append the pivot's detail rows directly into that predefined table before saving the file.
// Keywords: Aspose.Cells ShowDetail | C# PivotTable detail to ListObject | DetailTable parameter | populate existing Excel table | drill‑down pivot data | Aspose.Cells .NET example | Excel ListObject mapping
// Common Searches: Aspose.Cells ShowDetail map to existing table | C# map pivot detail rows into ListObject | Configure DetailTable parameter in Aspose.Cells | Fill predefined Excel table with pivot drill‑down data | Aspose.Cells pivot ShowDetail example
// Developer Intent: Insert the rows returned by PivotTable.ShowDetail into a pre‑created ListObject called DetailTable on another worksheet.
// Use Cases: Generate a drill‑down report where selecting a pivot category writes its underlying records into a styled table on a separate sheet. | Automate population of a template workbook that contains a formatted table with pivot detail data for scheduled reporting. | Export detailed sales or inventory records for a specific pivot item into a structured Excel table for downstream analysis.
// AI Prompts: Write C# code with Aspose.Cells that uses PivotTable.ShowDetail to fill an existing ListObject named DetailTable on a target worksheet. | Show how to configure the DetailTable parameter in ShowDetail to append rows to a predefined Excel table. | Explain the steps to activate the destination sheet before calling ShowDetail so the detail rows are written into the ListObject.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsDetailTableDemo
{
    // Demonstrates creating a workbook, adding a pivot table, defining a target sheet with an empty ListObject named DetailTable, and using PivotTable.ShowDetail to append the pivot's detail rows directly into that predefined table before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and get the first sheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                Cells dataCells = dataSheet.Cells;

                // -------------------------------------------------
                // 2. Populate sample data for the pivot table
                // -------------------------------------------------
                // Header row
                dataCells["A1"].PutValue("Category");
                dataCells["B1"].PutValue("Item");
                dataCells["C1"].PutValue("Quantity");

                // Data rows
                dataCells["A2"].PutValue("Fruit");
                dataCells["B2"].PutValue("Apple");
                dataCells["C2"].PutValue(10);

                dataCells["A3"].PutValue("Fruit");
                dataCells["B3"].PutValue("Banana");
                dataCells["C3"].PutValue(15);

                dataCells["A4"].PutValue("Vegetable");
                dataCells["B4"].PutValue("Carrot");
                dataCells["C4"].PutValue(20);

                dataCells["A5"].PutValue("Vegetable");
                dataCells["B5"].PutValue("Tomato");
                dataCells["C5"].PutValue(25);

                // -------------------------------------------------
                // 3. Add a pivot table based on the data range
                // -------------------------------------------------
                // The pivot will be placed starting at cell E3
                int pivotIndex = dataSheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivot = dataSheet.PivotTables[pivotIndex];

                // Row field: Category
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                // Data field: Sum of Quantity
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // -------------------------------------------------
                // 4. Create a target worksheet that already contains a predefined table
                // -------------------------------------------------
                Worksheet detailSheet = workbook.Worksheets.Add("DetailData");
                Cells detailCells = detailSheet.Cells;

                // Define a placeholder range for the table (A1:C1 will be the header)
                detailCells["A1"].PutValue("Item");
                detailCells["B1"].PutValue("Quantity");
                detailCells["C1"].PutValue("Category");

                // Add a ListObject (Excel table) that will receive the detail data
                // The table currently has only the header row; rows will be filled by ShowDetail
                int tableIndex = detailSheet.ListObjects.Add(0, 0, 0, 2, true);
                ListObject detailTable = detailSheet.ListObjects[tableIndex];
                detailTable.DisplayName = "DetailTable";

                // -------------------------------------------------
                // 5. Show detail for a specific pivot item and map it into the predefined table
                // -------------------------------------------------
                // Activate the detail sheet before calling ShowDetail
                workbook.Worksheets.ActiveSheetIndex = detailSheet.Index;

                // ShowDetail will append rows to the existing table starting at A2
                pivot.ShowDetail(
                    rowOffset: 1,          // first data row in the pivot's data region
                    columnOffset: 0,       // first data column in the pivot's data region
                    newSheet: false,       // place detail on the current (active) sheet
                    destRow: 1,            // start writing detail from row index 1 (A2)
                    destColumn: 0);        // start from column index 0 (A)

                // -------------------------------------------------
                // 6. Recalculate formulas (if any) and save the workbook
                // -------------------------------------------------
                workbook.CalculateFormula();

                string outputPath = "DetailTableMapped.xlsx";
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
