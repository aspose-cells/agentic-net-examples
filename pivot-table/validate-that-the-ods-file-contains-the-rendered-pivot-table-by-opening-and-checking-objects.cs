// Title: Validate that an ODS spreadsheet contains a rendered pivot table using Aspose.Cells for .NET
// AI Prompts: Load an ODS file with options that refresh pivot tables, then confirm the workbook includes at least one pivot table. | Write C# code to enumerate the pivot tables on the first worksheet and print the count of base fields for the first pivot table. | Add robust file‑existence verification and exception handling while checking for pivot tables in an ODS document.
// Common Searches: aspnet check if ODS file has pivot tables using Aspose.Cells | C# load ODS with OdsLoadOptions and verify pivot table rendering | how to programmatically confirm pivot table existence in OpenDocument spreadsheet | Aspose.Cells refresh pivot tables on load for ODS files example | detect missing pivot tables in ODS workbook with .NET
// Tags: initialize ODS workbook using OdsLoadOptions pivot refresh | list pivot tables on a worksheet with Aspose.Cells | display number of base fields for the first pivot table in C# | verify file existence before ODS pivot validation | enable classic pivot table mode when loading ODS

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example checks that a specified ODS file exists, loads it with OdsLoadOptions configured to refresh pivot tables and use classic mode, accesses the first worksheet, retrieves its PivotTableCollection, confirms at least one pivot table is present, and writes the base fields count of the first pivot table to the console while handling missing files and other exceptions.
    public class ValidateOdsPivotTable
    {
        public static void Run()
        {
            // Path to the ODS file to be validated
            string odsPath = "SamplePivot.ods";

            // Ensure the file exists to avoid FileNotFoundException
            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"File not found: {odsPath}");
                return;
            }

            try
            {
                // Configure ODS load options
                OdsLoadOptions loadOptions = new OdsLoadOptions
                {
                    // Refresh pivot tables while loading to ensure they are rendered
                    RefreshPivotTables = true,
                    // Load as classic pivot table if needed
                    IsClassicPivotTable = true
                };

                // Load the ODS workbook with the specified options
                Workbook workbook = new Workbook(odsPath, loadOptions);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the collection of pivot tables in the worksheet
                PivotTableCollection pivotTables = worksheet.PivotTables;

                // Validate that at least one pivot table exists
                if (pivotTables.Count > 0)
                {
                    // Get the first pivot table
                    PivotTable pivotTable = pivotTables[0];

                    // Output basic information to confirm rendering
                    Console.WriteLine("Pivot table found.");
                    Console.WriteLine("Base fields count: " + pivotTable.BaseFields.Count);
                    // RefreshDataFlag is obsolete; omitted per API guidance
                }
                else
                {
                    Console.WriteLine("No pivot tables were found in the ODS file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the ODS file:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateOdsPivotTable.Run();
        }
    }
}
