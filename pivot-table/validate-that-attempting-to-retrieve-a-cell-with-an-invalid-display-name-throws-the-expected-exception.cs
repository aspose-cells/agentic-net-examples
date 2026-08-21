// Title: Aspose.Cells .NET: GetCellByDisplayName throws exception for non‑existent pivot field
// Description: Creates a workbook with a pivot table, retrieves a cell using a valid data field display name, then calls GetCellByDisplayName with a missing display name, catches the resulting exception, and saves the file.
// Keywords: Aspose.Cells | C# | PivotTable | GetCellByDisplayName | invalid display name | exception handling | Aspose.Cells .NET | pivot field not found | cell retrieval error | Aspose.Cells API
// Common Searches: GetCellByDisplayName invalid field Aspose.Cells | What exception does GetCellByDisplayName throw when display name does not exist | How to catch missing pivot field error in Aspose.Cells | Aspose.Cells pivot table GetCellByDisplayName example | C# Aspose.Cells exception for non‑existent display name
// Developer Intent: Verify that GetCellByDisplayName raises the correct exception when the specified display name is not present in the pivot table.
// Use Cases: Validate error handling in applications that rely on pivot table cell lookup. | Unit‑test GetCellByDisplayName behavior for both valid and invalid display names. | Demonstrate safe workbook processing by catching missing field errors before saving. | Provide sample code for developers integrating Aspose.Cells pivot tables.
// AI Prompts: Generate an xUnit test that asserts GetCellByDisplayName throws ArgumentException for an unknown display name in an Aspose.Cells pivot table. | Show how to log the exception type and message when GetCellByDisplayName fails due to a missing field in C#. | Explain best practices for handling GetCellByDisplayName errors in Aspose.Cells .NET applications. | Create a reusable helper method that wraps GetCellByDisplayName with proper exception handling.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with a pivot table, retrieves a cell using a valid data field display name, then calls GetCellByDisplayName with a missing display name, catches the resulting exception, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Fruit";
        sheet.Cells["B1"].Value = "Quantity";
        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["B2"].Value = 10;
        sheet.Cells["A3"].Value = "Orange";
        sheet.Cells["B3"].Value = 15;
        sheet.Cells["A4"].Value = "Banana";
        sheet.Cells["B4"].Value = 20;

        // Add a pivot table to the worksheet
        int ptIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Configure the pivot table (row field and data field)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Demonstrate successful retrieval with a valid display name
        string validDisplayName = pivotTable.DataFields[0].DisplayName;
        Cell validCell = pivotTable.GetCellByDisplayName(validDisplayName);
        Console.WriteLine($"Valid display name \"{validDisplayName}\" returned cell {validCell.Name}");

        // Attempt to retrieve a cell using an invalid display name and verify exception
        string invalidDisplayName = "NonExistentField";
        try
        {
            // This call is expected to throw an exception because the display name does not exist
            Cell invalidCell = pivotTable.GetCellByDisplayName(invalidDisplayName);
            // If no exception is thrown, the test has failed
            Console.WriteLine("ERROR: No exception thrown for invalid display name.");
        }
        catch (Exception ex)
        {
            // Expected path: capture and display the exception type and message
            Console.WriteLine($"Caught expected exception: {ex.GetType().Name} - {ex.Message}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("PivotTable_InvalidDisplayName_Test.xlsx");
    }
}
