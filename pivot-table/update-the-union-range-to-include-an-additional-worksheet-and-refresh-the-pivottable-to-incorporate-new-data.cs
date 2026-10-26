// Title: How to extend a named union range with another worksheet and refresh all PivotTables using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a new worksheet to an existing named union range | C# refresh all PivotTables after changing the named range source in a workbook | Update RefersTo property to include another sheet and refresh pivot cache with Aspose.Cells | Programmatically extend DataRange union range with Sheet4 and refresh pivots | Aspose.Cells .NET modify union range and automatically refresh PivotTables
// Tags: named union range extension Aspose.Cells | worksheet addition to union range C# | pivot table cache refresh Aspose.Cells | update RefersTo property .NET | programmatic pivot refresh after range change

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads an existing workbook, checks for the named union range "DataRange", creates or retrieves a worksheet named "Sheet4", appends its $A$1:$D$100 area to the range's RefersTo string, updates the named range, iterates through every worksheet to call RefreshData on each PivotTable, and saves the workbook.
class UpdateUnionRangeAndRefreshPivot
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the existing union range (named range) to be updated
            string unionRangeName = "DataRange";

            // Get the existing named range object (if it exists)
            Name existingName = workbook.Worksheets.Names[unionRangeName];
            if (existingName == null)
            {
                Console.WriteLine($"Error: Named range '{unionRangeName}' does not exist.");
                return;
            }

            // Retrieve the current range reference (could be a union of multiple areas)
            string currentRef = existingName.RefersTo; // e.g., Sheet1!$A$1:$D$100,Sheet2!$A$1:$D$100

            // Define the additional worksheet and its data range to be added to the union
            string additionalSheetName = "Sheet4";
            string additionalRangeAddress = "$A$1:$D$100";

            // Ensure the additional worksheet exists; if not, create it
            Worksheet additionalSheet = workbook.Worksheets[additionalSheetName];
            if (additionalSheet == null)
            {
                int newIndex = workbook.Worksheets.Add();
                additionalSheet = workbook.Worksheets[newIndex];
                additionalSheet.Name = additionalSheetName;
                // (Optional) Populate the new sheet with sample data here if needed
            }

            // Build the reference string for the new range
            string newRefPart = $"'{additionalSheetName}'!{additionalRangeAddress}";

            // Combine the existing reference with the new part to form a union range
            string updatedRef = string.IsNullOrEmpty(currentRef) ? newRefPart : $"{currentRef},{newRefPart}";

            // Update the named range to refer to the new union range
            existingName.RefersTo = updatedRef;

            // Refresh all PivotTables in the workbook to incorporate the updated data source
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the data source of the PivotTable
                    pivotTable.RefreshData();
                    // No additional method required; RefreshData updates the cache.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
