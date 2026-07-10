using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace ValidateOdsPivotTable
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the ODS file to be validated
            string odsPath = "input.ods";

            // Create OdsLoadOptions and enable pivot table refresh on load
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            loadOptions.RefreshPivotTables = true; // Ensure pivot tables are rendered

            // Load the ODS workbook with the specified options
            Workbook workbook = new Workbook(odsPath, loadOptions);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the collection of pivot tables in the worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Validate that at least one pivot table exists
            if (pivotTables.Count > 0)
            {
                Console.WriteLine("Pivot table(s) found in the ODS file.");

                // Optionally, inspect the first pivot table for additional validation
                PivotTable pivot = pivotTables[0];

                // Check if the pivot table has any row fields (basic sanity check)
                if (pivot.RowFields.Count > 0)
                {
                    Console.WriteLine("Pivot table contains row fields, indicating it is rendered correctly.");
                }
                else
                {
                    Console.WriteLine("Pivot table exists but has no row fields.");
                }

                // Verify that the pivot data has been refreshed
                Console.WriteLine("RefreshDataOnOpeningFile flag: " + pivot.RefreshDataOnOpeningFile);
            }
            else
            {
                Console.WriteLine("No pivot tables were found in the ODS file.");
            }
        }
    }
}