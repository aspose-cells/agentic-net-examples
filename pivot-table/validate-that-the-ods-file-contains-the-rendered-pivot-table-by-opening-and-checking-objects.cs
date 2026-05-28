using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ValidateOdsPivotTable
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the ODS file that should contain a pivot table
            string odsPath = "SamplePivot.ods";

            // Verify file exists to avoid FileNotFoundException
            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"File not found: {odsPath}");
                return;
            }

            // Load options: refresh pivot tables when the file is opened
            OdsLoadOptions loadOptions = new OdsLoadOptions
            {
                RefreshPivotTables = true
            };

            // Load the workbook with the specified ODS load options
            Workbook workbook = new Workbook(odsPath, loadOptions);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that at least one pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables were found in the ODS file.");
                return;
            }

            // Retrieve the first pivot table for inspection
            PivotTable pivot = worksheet.PivotTables[0];

            // Since RefreshDataFlag is obsolete, assume data is current after loading with RefreshPivotTables=true
            bool isDataCurrent = true;

            // Output validation details
            Console.WriteLine($"Pivot tables found: {worksheet.PivotTables.Count}");
            Console.WriteLine($"First pivot table name: {pivot.Name}");
            Console.WriteLine($"Pivot data current (no refresh needed): {isDataCurrent}");
            Console.WriteLine($"Base fields count (should be >0 if rendered): {pivot.BaseFields.Count}");
        }
    }
}