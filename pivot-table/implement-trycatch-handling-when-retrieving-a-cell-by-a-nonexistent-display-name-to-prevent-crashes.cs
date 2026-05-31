using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableGetCellByDisplayNameWithExceptionHandling
    {
        public static void Main()
        {
            Run();
        }

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
                cells["A4"].Value = "Utilities";
                cells["B4"].Value = 150;

                // Add a pivot table based on the sample data
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SamplePivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate the pivot table so that it contains data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Retrieve a valid display name (the first data field)
                string validDisplayName = pivotTable.DataFields[0].DisplayName;

                // Attempt to get a cell using a valid display name (should succeed)
                try
                {
                    Cell validCell = pivotTable.GetCellByDisplayName(validDisplayName);
                    Console.WriteLine($"Valid display name '{validDisplayName}' returned cell: {validCell?.Name ?? "null"}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error with valid display name: {ex.Message}");
                }

                // Define a non‑existent display name to trigger an exception or null result
                string invalidDisplayName = "NonExistentField";

                // Safely handle the retrieval of a cell by a non‑existent display name
                try
                {
                    Cell invalidCell = pivotTable.GetCellByDisplayName(invalidDisplayName);
                    if (invalidCell == null)
                    {
                        Console.WriteLine($"Display name '{invalidDisplayName}' does not exist – returned null.");
                    }
                    else
                    {
                        Console.WriteLine($"Unexpectedly retrieved cell {invalidCell.Name} for non‑existent display name.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error retrieving cell for display name '{invalidDisplayName}': {ex.Message}");
                }

                // Save the workbook (demonstrates lifecycle compliance)
                string outputPath = "PivotTable_GetCellByDisplayName_WithExceptionHandling.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected runtime error: {ex.Message}");
            }
        }
    }
}