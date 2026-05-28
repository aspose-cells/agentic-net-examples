using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

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
        cells["B2"].Value = 100;
        cells["A3"].Value = "Drink";
        cells["B3"].Value = 50;

        // Add a pivot table covering the data range and place it at D5
        int pivotIndex = sheet.PivotTables.Add("A1:B3", "D5", "DemoPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add a row field and a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh and calculate the pivot table so that display names are generated
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Retrieve a valid display name (e.g., "Sum of Amount")
        string validDisplayName = pivotTable.DataFields[0].DisplayName;

        // Use GetCellByDisplayName with a valid name – should return a Cell object
        Cell validCell = pivotTable.GetCellByDisplayName(validDisplayName);
        Console.WriteLine($"Valid display name '{validDisplayName}' returned cell: {validCell?.Name}");

        // Define an invalid display name that does not exist in the pivot table
        string invalidDisplayName = "NonExistingField";

        // Attempt to retrieve a cell with the invalid display name and verify that an exception is thrown
        try
        {
            // This call is expected to fail
            Cell invalidCell = pivotTable.GetCellByDisplayName(invalidDisplayName);
            // If no exception occurs, the test has failed
            Console.WriteLine("ERROR: Expected exception was not thrown for invalid display name.");
        }
        catch (Exception ex)
        {
            // Expected path – output the exception type and message
            Console.WriteLine($"Expected exception caught: {ex.GetType().Name} - {ex.Message}");
        }

        // Save the workbook (optional, just to demonstrate normal lifecycle usage)
        workbook.Save("PivotTable_GetCellByDisplayName_ExceptionDemo.xlsx");
    }
}