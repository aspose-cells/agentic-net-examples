// Title: Validate Pivot Table Rendering in an ODS Workbook with Aspose.Cells for .NET
// Description: Loads an ODS file using OdsLoadOptions.RefreshPivotTables, accesses the first worksheet, checks the PivotTables collection, and confirms that the first pivot table contains base fields, indicating it is rendered.
// Keywords: Aspose.Cells | ODS | pivot table validation | RefreshPivotTables | C# | .NET workbook loading | PivotTable.BaseFields | check pivot existence | OdsLoadOptions example | Aspose.Cells pivot
// Common Searches: how to verify a pivot table in an ODS file with Aspose.Cells | Aspose.Cells C# refresh pivot tables on load | detect pivot tables in ODS workbook using .NET | check base fields of a pivot table Aspose.Cells | validate rendered pivot table in ODS with Aspose
// Developer Intent: Confirm that an ODS workbook contains a rendered pivot table by loading the file and inspecting its pivot table objects.
// Use Cases: Load an ODS workbook with RefreshPivotTables enabled and determine if any pivot tables are present. | Retrieve the first worksheet’s PivotTables collection to verify the existence of a pivot table. | Examine the BaseFields count of a PivotTable to ensure the table has been rendered. | Handle missing files or loading errors gracefully while performing the validation.
// AI Prompts: Write C# code that opens an ODS workbook with Aspose.Cells, refreshes pivot tables on load, and returns true when a pivot table with at least one base field is found. | Show how to catch FileNotFoundException and other errors when validating pivot tables in an ODS file using Aspose.Cells. | Explain the relationship between OdsLoadOptions.RefreshPivotTables and PivotTable.BaseFields for confirming that a pivot table is rendered.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an ODS file using OdsLoadOptions.RefreshPivotTables, accesses the first worksheet, checks the PivotTables collection, and confirms that the first pivot table contains base fields, indicating it is rendered.
    public class ValidateOdsPivotTable
    {
        public static void Run()
        {
            // Path to the ODS file that should contain a pivot table
            string odsPath = "SamplePivot.ods";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"Error: The file \"{odsPath}\" was not found.");
                return;
            }

            try
            {
                // Configure load options to refresh pivot tables when the file is opened
                OdsLoadOptions loadOptions = new OdsLoadOptions
                {
                    RefreshPivotTables = true
                };

                // Load the ODS workbook with the specified options
                Workbook workbook = new Workbook(odsPath, loadOptions);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Check whether any pivot tables are present
                if (worksheet.PivotTables.Count > 0)
                {
                    // Get the first pivot table
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Validate that the pivot table has at least one base field (i.e., it is rendered)
                    bool hasBaseFields = pivotTable.BaseFields.Count > 0;

                    // Since RefreshDataOnOpen is not available, rely on the load option used
                    bool isRefreshed = loadOptions.RefreshPivotTables;

                    Console.WriteLine("Pivot table validation result:");
                    Console.WriteLine($"- Pivot table exists: true");
                    Console.WriteLine($"- Has base fields: {hasBaseFields}");
                    Console.WriteLine($"- RefreshPivotTables option set: {isRefreshed}");
                }
                else
                {
                    Console.WriteLine("Pivot table validation result:");
                    Console.WriteLine("- Pivot table exists: false");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateOdsPivotTable.Run();
        }
    }
}
