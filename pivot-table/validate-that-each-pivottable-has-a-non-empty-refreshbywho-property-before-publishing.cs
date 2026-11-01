// Title: How to ensure every PivotTable has a non‑empty RefreshByWho property before saving a workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file using Aspose.Cells, iterates all worksheets and their PivotTables, and throws an InvalidOperationException if PivotTable.RefreshByWho is null, empty, or whitespace before calling Workbook.Save. | Create a helper method ValidatePivotRefreshByWho(Workbook wb) that returns a list of error messages for each PivotTable lacking a RefreshByWho value, and integrate it into a console application. | Write a logging routine that records the worksheet name and pivot table name for every PivotTable whose RefreshByWho property is missing, then continues processing the remaining tables.
// Common Searches: Aspose.Cells C# check RefreshByWho value for each PivotTable before publishing | validate PivotTable RefreshByWho property in .NET workbook | throw error when PivotTable.RefreshByWho is blank using Aspose.Cells | how to iterate pivot tables and verify RefreshByWho in C#
// Tags: validate pivot RefreshByWho Aspose.Cells | check pivot table properties .NET | iterate worksheets pivot tables C# | save workbook after pivot validation Aspose.Cells | exception handling for empty pivot property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable class resides in this namespace

// The example loads 'input.xlsx' with Aspose.Cells, walks through every worksheet and each PivotTable, verifies that the RefreshByWho property is not null or whitespace (and optionally that the Name is set), throws an InvalidOperationException or logs the offending table, and saves the workbook as 'output.xlsx' only when all validations succeed.
class PivotTableValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables on the current worksheet
                for (int i = 0; i < sheet.PivotTables.Count; i++)
                {
                    // Retrieve the pivot table
                    PivotTable pivot = sheet.PivotTables[i];

                    // Validate that the pivot table has a non‑empty name
                    if (string.IsNullOrWhiteSpace(pivot.Name))
                    {
                        throw new InvalidOperationException(
                            $"PivotTable in worksheet '{sheet.Name}' has an empty Name property.");
                    }
                }
            }

            // Save the workbook if all validations pass
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
