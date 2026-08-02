// Title: Create a Dynamic Pivot Table in Aspose.Cells for .NET Using a ListObject (Table)
// Description: Demonstrates how to build a workbook, define a ListObject that auto‑expands, and generate a pivot table whose source is the table name. The example adds row and data fields, refreshes the pivot, and saves the file as DynamicPivotTable.xlsx.
// Keywords: Aspose.Cells | C# | .NET | dynamic pivot table | ListObject | Excel table source | auto‑expand range | pivot refresh | programmatic pivot creation | Excel reporting
// Common Searches: Aspose.Cells dynamic pivot source | C# create pivot from ListObject | expandable pivot range Aspose | refresh pivot after adding rows .NET | use Excel table as pivot source Aspose.Cells
// Developer Intent: Generate a pivot table that automatically includes new rows added to the source worksheet without redefining the range.
// Use Cases: Sales dashboard that updates when daily transactions are appended. | Financial workbook where accountants add month‑end entries and the summary pivot adjusts instantly. | Self‑service reporting tool allowing users to extend data tables and see refreshed pivot results.
// AI Prompts: Show code to add rows to the ListObject and refresh the associated pivot table in Aspose.Cells. | Explain how to bind a pivot table to a named Excel table and keep it synchronized as the table grows. | Provide an example of configuring row and data fields after the source table expands using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsDynamicPivot
{
    // Demonstrates how to build a workbook, define a ListObject that auto‑expands, and generate a pivot table whose source is the table name. The example adds row and data fields, refreshes the pivot, and saves the file as DynamicPivotTable.xlsx.
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
                dataSheet.Name = "SourceData";

                // Add header
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");

                // Add some initial rows
                dataSheet.Cells["A2"].PutValue("A");
                dataSheet.Cells["B2"].PutValue(10);
                dataSheet.Cells["A3"].PutValue("B");
                dataSheet.Cells["B3"].PutValue(20);
                dataSheet.Cells["A4"].PutValue("A");
                dataSheet.Cells["B4"].PutValue(30);

                // -------------------------------------------------
                // 2. Convert the data range into a Table (ListObject)
                //    Tables automatically expand when new rows are added,
                //    which makes the pivot source dynamic.
                // -------------------------------------------------
                // Define the initial data range (including headers)
                // Rows and columns are zero‑based, so A1:B5 corresponds to rows 0‑4 and columns 0‑1
                string tableSource = "A1:B5";

                // Add the ListObject (table) to the worksheet with a name
                int tableIndex = dataSheet.ListObjects.Add("DynamicTable", tableSource, true);
                ListObject table = dataSheet.ListObjects[tableIndex];

                // -------------------------------------------------
                // 3. Add a new worksheet for the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // -------------------------------------------------
                // 4. Add the pivot table using the table name as source data
                // -------------------------------------------------
                string sourceData = "DynamicTable"; // Table name as source
                string destCell = "A1";              // Upper‑left corner of the pivot table
                string pivotName = "MyDynamicPivot";

                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, destCell, pivotName);
                PivotTable pivotTable = pivotTables[pivotIndex];

                // -------------------------------------------------
                // 5. Configure the pivot fields
                // -------------------------------------------------
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // -------------------------------------------------
                // 6. Refresh the pivot to ensure it reflects the current data
                // -------------------------------------------------
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 7. Save the workbook
                // -------------------------------------------------
                workbook.Save("DynamicPivotTable.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
