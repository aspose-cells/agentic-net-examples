// Title: Aspose.Cells for .NET – Verify PivotTable.GetCellByDisplayName throws exception for an invalid display name
// Description: This example creates a workbook, adds a pivot table, defines row and data fields, refreshes the table, and then calls PivotTable.GetCellByDisplayName with a non‑existent field name. The code catches the expected exception and logs its type and message, demonstrating proper error handling for invalid pivot field references.
// Keywords: Aspose.Cells | PivotTable | GetCellByDisplayName | invalid display name | exception handling | .NET | Aspose.Cells for .NET | pivot table error | catch exception
// Common Searches: Aspose.Cells GetCellByDisplayName invalid field exception | how to test PivotTable.GetCellByDisplayName throws error | exception type for unknown display name in Aspose.Cells pivot table | validate error handling for PivotTable.GetCellByDisplayName .NET
// Developer Intent: Confirm that calling PivotTable.GetCellByDisplayName with a non‑existent display name raises the appropriate exception.
// Use Cases: Create a unit test that asserts an exception is thrown for an unknown pivot field name. | Wrap GetCellByDisplayName in try‑catch to log exception details during debugging. | Validate pivot table configuration by ensuring all required display names exist before accessing cells.
// AI Prompts: Generate an MSTest method that verifies PivotTable.GetCellByDisplayName throws a specific exception for an unknown display name. | Provide sample code that catches the exception from GetCellByDisplayName and logs its type and message. | Explain which exception Aspose.Cells throws when GetCellByDisplayName receives an invalid display name and recommend handling strategies.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotValidation
{
    // This example creates a workbook, adds a pivot table, defines row and data fields, refreshes the table, and then calls PivotTable.GetCellByDisplayName with a non‑existent field name. The code catches the expected exception and logs its type and message, demonstrating proper error handling for invalid pivot field references.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Drink";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Snack";
            cells["B4"].Value = 45;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row field and data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table so that it is ready for queries
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Attempt to retrieve a cell using an invalid display name.
            // The display name "NonExistentField" does not exist in the pivot table.
            string invalidDisplayName = "NonExistentField";

            try
            {
                // This call is expected to throw an exception because the display name is invalid.
                Cell cell = pivotTable.GetCellByDisplayName(invalidDisplayName);

                // If no exception is thrown, the test has failed.
                Console.WriteLine("Test Failed: No exception was thrown for an invalid display name.");
            }
            catch (Exception ex)
            {
                // Expected path: an exception should be thrown.
                // Output the exception type and message for verification.
                Console.WriteLine("Test Passed: Caught expected exception.");
                Console.WriteLine($"Exception Type: {ex.GetType().FullName}");
                Console.WriteLine($"Message: {ex.Message}");
            }

            // Optionally save the workbook (not required for the validation test)
            // workbook.Save("PivotTable_InvalidDisplayName_Test.xlsx");
        }
    }
}
