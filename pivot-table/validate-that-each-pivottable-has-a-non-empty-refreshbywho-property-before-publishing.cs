// Title: Validate non‑empty RefreshedByWho for every PivotTable before publishing with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel workbook, iterates all worksheets, refreshes each PivotTable whose RefreshedByWho property is empty, checks the property, and throws an InvalidOperationException if it remains blank. | Modify the sample program to log the names of PivotTables that have an empty RefreshedByWho value to a text file and continue processing the remaining tables instead of aborting. | Create a reusable C# method named EnsurePivotRefreshInfo that accepts a Workbook instance, refreshes any PivotTable lacking RefreshedByWho, validates the field, and returns a list of tables still missing the information.
// Common Searches: Aspose.Cells C# how to verify RefreshedByWho property of pivot tables before saving | check if pivot table refreshed by user is set using Aspose.Cells .NET | C# iterate worksheets and validate pivot table RefreshedByWho value | throw exception when RefreshedByWho is empty in Aspose.Cells pivot tables | refresh pivot tables programmatically with Aspose.Cells to populate RefreshedByWho
// Tags: Aspose.Cells PivotTable RefreshedByWho validation | C# refresh pivot tables programmatically | Excel workbook pivot table property check | Aspose.Cells throw exception on missing RefreshedByWho | automated pivot table refresh before workbook publish

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableValidation
{
    // Loads an Excel workbook, walks through each worksheet's PivotTable collection, refreshes any pivot table with an empty RefreshedByWho property, validates that the property is populated, throws an InvalidOperationException if validation fails, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (uses the standard load rule)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of pivot tables in the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Iterate through each pivot table
                foreach (PivotTable pt in pivots)
                {
                    // Ensure the pivot table has been refreshed so that RefreshedByWho is populated
                    // If the workbook was created without a refresh, call RefreshData and CalculateData
                    if (string.IsNullOrEmpty(pt.RefreshedByWho))
                    {
                        // Refresh the pivot table data
                        pt.RefreshData();
                        pt.CalculateData();
                    }

                    // Validate that RefreshedByWho is not empty after refresh
                    if (string.IsNullOrEmpty(pt.RefreshedByWho))
                    {
                        // Throw an exception or handle the validation failure as needed
                        throw new InvalidOperationException(
                            $"PivotTable '{pt.Name}' in worksheet '{sheet.Name}' does not have a non‑empty RefreshedByWho property.");
                    }
                }
            }

            // Save the workbook after successful validation (uses the standard save rule)
            string outputPath = "ValidatedWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}
