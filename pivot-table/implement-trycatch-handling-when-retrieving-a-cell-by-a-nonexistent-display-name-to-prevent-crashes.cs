// Title: C# – Safe Retrieval of PivotTable Cell by Display Name with Exception Handling in Aspose.Cells
// Description: Shows how to build a workbook, add a pivot table, and fetch cells via PivotTable.GetCellByDisplayName. Includes try‑catch handling for a missing display name, null checks, error logging, and workbook saving to prevent crashes.
// Keywords: Aspose.Cells | C# | PivotTable | GetCellByDisplayName | exception handling | invalid display name | null check | pivot cell retrieval | error logging | Excel automation
// Common Searches: Aspose.Cells GetCellByDisplayName try catch | handle missing pivot field in Aspose.Cells | C# retrieve pivot table cell by display name safely | Aspose.Cells exception when display name not found | prevent crash GetCellByDisplayName Aspose
// Developer Intent: Add robust error handling around PivotTable.GetCellByDisplayName so the program continues running when the specified display name is absent.
// Use Cases: Validate user‑provided pivot field names before accessing data. | Log detailed errors for unsupported display names in reporting pipelines. | Supply fallback values when a required pivot field is missing. | Keep batch Excel processing jobs stable despite configuration issues. | Debug pivot table setups during development without terminating the app.
// AI Prompts: Generate C# code that wraps PivotTable.GetCellByDisplayName in a helper method with try‑catch and returns null on failure. | Explain step‑by‑step how to add exception handling for GetCellByDisplayName in Aspose.Cells. | Create a logging utility for Aspose.Cells pivot table errors when a display name does not exist. | Write a unit test that verifies GetCellByDisplayName throws an exception for an invalid field and is caught properly.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, add a pivot table, and fetch cells via PivotTable.GetCellByDisplayName. Includes try‑catch handling for a missing display name, null checks, error logging, and workbook saving to prevent crashes.
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
                cells["A4"].Value = "Food";
                cells["B4"].Value = 150;
                cells["A5"].Value = "Travel";
                cells["B5"].Value = 200;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: add row field and data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate the pivot table so that display names are generated
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Retrieve a valid display name (for demonstration)
                string validDisplayName = pivotTable.DataFields[0].DisplayName;
                Console.WriteLine($"Valid display name: {validDisplayName}");

                // Attempt to retrieve a cell using a non‑existent display name
                string invalidDisplayName = "NonExistentField";
                Cell cell = null;

                try
                {
                    // This call will throw if the display name does not exist
                    cell = pivotTable.GetCellByDisplayName(invalidDisplayName);
                    // If no exception, still check for null (method may return null in some versions)
                    if (cell == null)
                    {
                        Console.WriteLine($"Display name '{invalidDisplayName}' returned null.");
                    }
                    else
                    {
                        Console.WriteLine($"Cell found: {cell.Name} with value '{cell.Value}'.");
                    }
                }
                catch (Exception ex)
                {
                    // Handle the error gracefully without crashing the application
                    Console.WriteLine($"Error retrieving cell for display name '{invalidDisplayName}': {ex.Message}");
                }

                // Demonstrate successful retrieval using the valid display name
                try
                {
                    Cell validCell = pivotTable.GetCellByDisplayName(validDisplayName);
                    Console.WriteLine($"Valid cell: {validCell.Name}, Value: {validCell.Value}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error with valid display name: {ex.Message}");
                }

                // Save the workbook (output file name can be adjusted as needed)
                string outputPath = "PivotTable_GetCellByDisplayName_WithExceptionHandling.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
