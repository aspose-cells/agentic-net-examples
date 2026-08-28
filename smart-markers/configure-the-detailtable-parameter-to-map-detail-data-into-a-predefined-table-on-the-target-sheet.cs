// Title: Map pivot table detail rows into a predefined ListObject (DetailTable) on a different worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses PivotTable.ShowDetail to extract detail rows from a pivot table and writes them into an existing ListObject on another sheet. | Create a reusable method that accepts a PivotTable object and a ListObject reference, copies the extracted detail data, and populates the ListObject columns. | Show how to copy the range returned by ShowDetail into a target worksheet and bind it to a ListObject named DetailTable.
// Common Searches: aspocells showdetail copy pivot detail to another worksheet listobject c# | how to populate an existing Excel table with pivot item details using Aspose.Cells | C# extract pivot table detail rows and insert into predefined ListObject | map pivot table ShowDetail output to a ListObject in Aspose.Cells for .NET | transfer pivot detail data to a separate sheet table programmatically
// Tags: pivot table ShowDetail to ListObject mapping | Aspose.Cells populate existing ListObject from pivot detail | C# extract pivot detail rows into Excel table | copy pivot detail rows to another worksheet with Aspose.Cells | predefined ListObject data insertion using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsDetailTableDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, then creates a second worksheet with an empty ListObject named DetailTable. Using PivotTable.ShowDetail it extracts the detail rows of the first data cell, copies the resulting range to the target sheet, and writes each Category and Amount into the ListObject rows before saving the file as DetailTableMapped.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data for the pivot table
                // -------------------------------------------------
                Worksheet sourceSheet = workbook.Worksheets[0];
                Cells srcCells = sourceSheet.Cells;

                // Header
                srcCells["A1"].PutValue("Category");
                srcCells["B1"].PutValue("Amount");

                // Sample rows
                srcCells["A2"].PutValue("Food");
                srcCells["B2"].PutValue(120);
                srcCells["A3"].PutValue("Food");
                srcCells["B3"].PutValue(80);
                srcCells["A4"].PutValue("Travel");
                srcCells["B4"].PutValue(200);
                srcCells["A5"].PutValue("Travel");
                srcCells["B5"].PutValue(150);
                srcCells["A6"].PutValue("Utilities");
                srcCells["B6"].PutValue(90);

                // -------------------------------------------------
                // 2. Create a pivot table based on the source data
                // -------------------------------------------------
                // Data range: A1:B6
                int pivotIndex = sourceSheet.PivotTables.Add("A1:B6", "E3", "SalesPivot");
                PivotTable pivot = sourceSheet.PivotTables[pivotIndex];

                // Add fields: Category as row, Amount as data (sum)
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate the pivot to ensure data is ready
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 3. Prepare the target sheet with a predefined table
                // -------------------------------------------------
                Worksheet targetSheet = workbook.Worksheets.Add("DetailTableSheet");
                Cells tgtCells = targetSheet.Cells;

                // Create an empty table (ListObject) at A1:B1 (header row only)
                tgtCells["A1"].PutValue("Category");
                tgtCells["B1"].PutValue("Amount");
                int tableIndex = targetSheet.ListObjects.Add(0, 0, 0, 1, true);
                ListObject detailTable = targetSheet.ListObjects[tableIndex];
                detailTable.DisplayName = "DetailTable";

                // -------------------------------------------------
                // 4. Map the detail data of a pivot item into the predefined table
                // -------------------------------------------------
                // ShowDetail extracts detail rows for the first data cell of the pivot
                // and writes them to the same worksheet (sourceSheet) starting at row 1, column 0.
                // After extraction, copy the data into the predefined table on targetSheet.
                pivot.ShowDetail(
                    rowOffset: 1,
                    columnOffset: 0,
                    newSheet: false,
                    destRow: 1,
                    destColumn: 0);

                // Copy extracted data (starting at E4 in sourceSheet) to the target table
                // Determine the range of extracted data
                int startRow = 3; // zero‑based index for row 4 (E4)
                int startCol = 4; // zero‑based index for column E
                int lastRow = sourceSheet.Cells.MaxDataRow;
                int lastCol = sourceSheet.Cells.MaxDataColumn;

                // Ensure there is data to copy
                if (lastRow >= startRow)
                {
                    // Copy each row into the target table starting after its header
                    int targetRow = 1; // zero‑based index for row 2 in target sheet
                    for (int r = startRow; r <= lastRow; r++)
                    {
                        // Category (column A) -> source column E (index 4)
                        object category = sourceSheet.Cells[r, startCol].Value;
                        // Amount (column B) -> source column F (index 5) if exists
                        object amount = (startCol + 1 <= lastCol) ? sourceSheet.Cells[r, startCol + 1].Value : null;

                        tgtCells[targetRow, 0].PutValue(category);
                        tgtCells[targetRow, 1].PutValue(amount);
                        targetRow++;
                    }
                }

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("DetailTableMapped.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
