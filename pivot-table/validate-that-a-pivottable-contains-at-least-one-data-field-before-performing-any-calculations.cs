// Title: Validate PivotTable DataFields before calculation with Aspose.Cells for .NET (C#)
// Description: This C# example shows how to create a workbook, populate it with date and sales data, add a PivotTable, verify that the PivotTable has at least one data field, automatically insert a default "Sales" field when none exist, refresh the pivot cache, recalculate the pivot values, and save the file.
// Keywords: Aspose.Cells | PivotTable validation | C# | .NET | DataFields count | Add default data field | RefreshData | CalculateData | Excel automation | pivot cache
// Common Searches: Aspose.Cells check PivotTable data fields C# | Add default data field to PivotTable Aspose | Validate PivotTable before CalculateData | RefreshData and CalculateData Aspose.Cells example | C# code to ensure PivotTable has a data field
// Developer Intent: Confirm a PivotTable contains at least one data field before invoking RefreshData and CalculateData.
// Use Cases: Prevent runtime errors by automatically adding a numeric field when a newly created PivotTable lacks data fields. | Integrate a validation step in automated report generation to guarantee pivot calculations succeed. | Log a warning and insert a fallback data field in dynamic workbook workflows that build PivotTables on the fly.
// AI Prompts: Generate C# code using Aspose.Cells that checks PivotTable.DataFields.Count and adds a specified field if the collection is empty. | Write a method that logs a message and inserts a default "Sales" data field before calling CalculateData on a PivotTable. | Provide an example of safely refreshing and calculating a PivotTable after ensuring at least one data field is present.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# example shows how to create a workbook, populate it with date and sales data, add a PivotTable, verify that the PivotTable has at least one data field, automatically insert a default "Sales" field when none exist, refresh the pivot cache, recalculate the pivot values, and save the file.
    public class PivotTableDataFieldValidationDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample source data for the pivot table
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["B1"].PutValue("Sales");
            DateTime baseDate = new DateTime(2022, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                worksheet.Cells[$"A{i + 2}"].PutValue(baseDate.AddDays(i));
                worksheet.Cells[$"B{i + 2}"].PutValue(100 + i * 20);
            }

            // Add a pivot table based on the source data
            int pivotIndex = worksheet.PivotTables.Add("A1:B11", "D3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add a row field (Date)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // ---- Validation: ensure at least one data field exists before calculations ----
            // If no data fields are present, add a default one (Sales)
            if (pivotTable.DataFields.Count == 0)
            {
                // Add the "Sales" column as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                Console.WriteLine("Data field was missing; added default data field 'Sales'.");
            }

            // Refresh the pivot cache and calculate the pivot data
            pivotTable.RefreshData();      // Correct method to refresh cache
            pivotTable.CalculateData();   // Recalculate pivot values

            // Save the workbook
            workbook.Save("PivotTableDataFieldValidationDemo.xlsx");
        }
    }
}
