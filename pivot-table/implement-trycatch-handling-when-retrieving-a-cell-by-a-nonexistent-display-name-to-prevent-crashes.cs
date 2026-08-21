// Title: C# Aspose.Cells PivotTable GetCellByDisplayName with Exception Handling for Missing Fields
// Description: Demonstrates how to safely retrieve a pivot table cell by its display name using Aspose.Cells. The example creates a workbook, builds a pivot table, fetches a cell for a valid field, and gracefully handles a non‑existent display name with try‑catch blocks and null checks, ensuring the application continues and the file is saved.
// Keywords: Aspose.Cells PivotTable GetCellByDisplayName | C# exception handling Aspose.Cells | handle missing pivot field | prevent crash GetCellByDisplayName | pivot table cell retrieval safe | Aspose.Cells try catch example
// Common Searches: Aspose.Cells GetCellByDisplayName exception | C# pivot table retrieve cell by display name safely | how to catch error when display name not found Aspose.Cells | prevent crash when GetCellByDisplayName fails
// Developer Intent: Add robust error handling around PivotTable.GetCellByDisplayName so that missing or incorrect display names do not terminate the program.
// Use Cases: Fetch the address of a valid data field cell in a pivot table. | Attempt to access a cell with an invalid display name and log a friendly message instead of crashing. | Continue processing and save the workbook even when one or more display names are incorrect.
// AI Prompts: Generate C# code that creates a pivot table with Aspose.Cells and retrieves cells by display name using try‑catch to handle missing fields. | Write a helper method that wraps PivotTable.GetCellByDisplayName, returns null on failure, and logs the exception message. | Show how to loop through a list of display names, call GetCellByDisplayName for each, and handle exceptions individually without stopping the loop.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to safely retrieve a pivot table cell by its display name using Aspose.Cells. The example creates a workbook, builds a pivot table, fetches a cell for a valid field, and gracefully handles a non‑existent display name with try‑catch blocks and null checks, ensuring the application continues and the file is saved.
    public class PivotTableGetCellByDisplayNameWithExceptionHandling
    {
        public static void Run()
        {
            try
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
                cells["A3"].Value = "Travel";
                cells["B3"].Value = 80;
                cells["A4"].Value = "Supplies";
                cells["B4"].Value = 45;

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SamplePivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (row field and data field)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh data and calculate the pivot table
                pivotTable.RefreshData();      // Correct API to refresh the cache
                pivotTable.CalculateData();

                // Retrieve an existing display name (for demonstration)
                string existingDisplayName = pivotTable.DataFields[0].DisplayName;

                // Attempt to get a cell by a valid display name
                try
                {
                    Cell validCell = pivotTable.GetCellByDisplayName(existingDisplayName);
                    Console.WriteLine($"Valid cell retrieved: {validCell?.Name ?? "null"}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving valid cell: {ex.Message}");
                }

                // Attempt to get a cell by a non‑existent display name
                string nonExistentDisplayName = "NonExistentField";

                try
                {
                    // This call may throw an exception if the display name does not exist
                    Cell invalidCell = pivotTable.GetCellByDisplayName(nonExistentDisplayName);
                    // If no exception is thrown, still check for null
                    if (invalidCell == null)
                    {
                        Console.WriteLine($"Display name \"{nonExistentDisplayName}\" does not correspond to any cell (null returned).");
                    }
                    else
                    {
                        Console.WriteLine($"Unexpected cell retrieved: {invalidCell.Name}");
                    }
                }
                catch (Exception ex)
                {
                    // Handle the exception gracefully to prevent the application from crashing
                    Console.WriteLine($"Handled exception for display name \"{nonExistentDisplayName}\": {ex.Message}");
                }

                // Save the workbook (output file)
                workbook.Save("PivotTable_GetCellByDisplayName_WithExceptionHandling.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Entry point required for compilation
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
