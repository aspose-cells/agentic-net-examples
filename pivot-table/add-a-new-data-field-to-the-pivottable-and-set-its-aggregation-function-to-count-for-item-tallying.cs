// Title: C# – Add a Count Data Field to an Aspose.Cells PivotTable
// Description: Demonstrates how to create a workbook with sample data, add a PivotTable, place "Category" in the row area, add "Item" as a data field, set its aggregation to Count using ConsolidationFunction.Count, refresh and calculate the pivot, and save the result as an Excel file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells PivotTable Count | C# add data field Count | ConsolidationFunction.Count Aspose | Aspose.Cells .NET pivot aggregation | Excel PivotTable count items C# | PivotFieldType.Data Aspose.Cells | RefreshData CalculateData Aspose
// Common Searches: Aspose.Cells set pivot data field to Count C# | How to add Count aggregation to a PivotTable using Aspose.Cells | C# code for counting items in an Aspose.Cells pivot table | Aspose.Cells PivotTable Count function example | Add data field and set ConsolidationFunction.Count in .NET
// Developer Intent: Add a new data field to a PivotTable and configure it to use the Count aggregation function.
// Use Cases: Count occurrences of each item within categories for inventory reports. | Create a frequency summary of products sold per category in sales dashboards. | Generate a quick tally of record counts for data validation or audit logs.
// AI Prompts: Generate C# code with Aspose.Cells that adds an "Item" data field to a PivotTable and sets its function to Count. | Explain how to change the aggregation of an existing Aspose.Cells PivotTable data field to Count and recalculate the table. | Provide step‑by‑step instructions to refresh and calculate a PivotTable after adding a Count data field using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // Demonstrates how to create a workbook with sample data, add a PivotTable, place "Category" in the row area, add "Item" as a data field, set its aggregation to Count using ConsolidationFunction.Count, refresh and calculate the pivot, and save the result as an Excel file with Aspose.Cells for .NET.
    class AddDataFieldWithCount
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample source data
            // Column A: Category, Column B: Item (to be counted)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue("Apple");
            cells["A3"].PutValue("Fruit");
            cells["B3"].PutValue("Orange");
            cells["A4"].PutValue("Vegetable");
            cells["B4"].PutValue("Carrot");
            cells["A5"].PutValue("Fruit");
            cells["B5"].PutValue("Apple");
            cells["A6"].PutValue("Vegetable");
            cells["B6"].PutValue("Broccoli");

            // Add a pivot table based on the source range A1:B6, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add "Category" as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add "Item" as a data field
            int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Item");

            // Set the aggregation function of the newly added data field to Count
            PivotField dataField = pivotTable.DataFields[dataFieldPos];
            dataField.Function = ConsolidationFunction.Count;

            // Refresh and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_CountField.xlsx");
        }
    }
}
