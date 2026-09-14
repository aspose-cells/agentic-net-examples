// Title: Safely retrieve a pivot table cell by display name with try‑catch handling for missing fields using Aspose.Cells for .NET
// AI Prompts: Write C# code that calls PivotTable.GetCellByDisplayName inside a try‑catch block and logs a warning when the display name is not found. | Generate a method that attempts to get a pivot cell by its display name, checks for a null result, and returns a default value without throwing. | Show how to wrap pivot table data refresh and cell retrieval in exception handling to prevent application crashes in Aspose.Cells.
// Common Searches: Aspose.Cells how to handle exception when GetCellByDisplayName returns null for invalid pivot field | C# example of try catch around PivotTable.GetCellByDisplayName for non‑existent display name | prevent crash retrieving pivot cell by display name Aspose.Cells .NET | check for missing pivot field before calling GetCellByDisplayName in Aspose.Cells
// Tags: Aspose.Cells PivotTable GetCellByDisplayName error handling | C# try-catch Aspose.Cells pivot table | null check PivotTable cell retrieval | handling missing display name Aspose.Cells | pivot table cell access Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook with a pivot table, refreshes its data, and demonstrates safe retrieval of cells by display name using try‑catch blocks and null checks to handle both valid and non‑existent field names before saving the file.
public class PivotTableGetCellByDisplayNameDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
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
        cells["A3"].Value = "Travel";
        cells["B3"].Value = 200;

        // Add a pivot table based on the sample data
        int ptIndex = sheet.PivotTables.Add("A1:B3", "D5", "Pivot1");
        PivotTable pivot = sheet.PivotTables[ptIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh pivot cache and calculate data (correct API usage)
        try
        {
            pivot.RefreshData();
            pivot.CalculateData();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error refreshing pivot data: {ex.Message}");
        }

        // Retrieve a cell using a valid display name (should succeed)
        string validDisplayName = pivot.DataFields[0].DisplayName;
        Cell validCell = pivot.GetCellByDisplayName(validDisplayName);
        Console.WriteLine($"Valid cell name: {validCell?.Name ?? "null"}");

        // Attempt to retrieve a cell using a non‑existent display name
        string invalidDisplayName = "NonExistentField";

        try
        {
            Cell invalidCell = pivot.GetCellByDisplayName(invalidDisplayName);
            if (invalidCell == null)
            {
                Console.WriteLine($"Cell not found for display name: {invalidDisplayName}");
            }
            else
            {
                // This path is unexpected but handled safely
                Console.WriteLine($"Unexpected cell returned: {invalidCell.Name}");
            }
        }
        catch (Exception ex)
        {
            // Gracefully handle any exception thrown by GetCellByDisplayName
            Console.WriteLine($"Error retrieving cell by display name '{invalidDisplayName}': {ex.Message}");
        }

        // Save the workbook
        try
        {
            workbook.Save("PivotTable_GetCellByDisplayName_Demo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
