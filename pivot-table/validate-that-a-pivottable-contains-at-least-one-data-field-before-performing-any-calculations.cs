// Title: How to validate a PivotTable has at least one data field before calling RefreshData with Aspose.Cells for .NET
// AI Prompts: Generate C# code that checks PivotTable.DataFields.Count and throws a descriptive exception if the count is zero before invoking RefreshData using Aspose.Cells. | Create a .NET method that loads an Excel workbook, ensures the first worksheet's PivotTable contains at least one value field, and only then calls PivotTable.RefreshData. | Write error‑handling logic for Aspose.Cells that logs a clear message and skips RefreshData when a PivotTable lacks data fields.
// Common Searches: Aspose.Cells C# verify PivotTable contains a value field before RefreshData | Check if PivotTable.DataFields is empty in .NET Excel processing | Prevent RefreshData error when PivotTable has no data fields using Aspose.Cells | C# code sample to validate PivotTable data fields count with Aspose.Cells | How to handle missing data fields in a PivotTable with Aspose.Cells for .NET
// Tags: aspose.cells pivot-table datafield check | c# refreshdata after pivot validation | excel workbook pivot-table empty datafield handling | aspose.cells verify pivot-table value field | c# validate pivot-table before calculations

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook, confirms the first worksheet contains a PivotTable with at least one data field, refreshes the PivotTable only after this validation, and saves the file, providing clear error messages for missing files, absent PivotTables, or empty data fields.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTables found on the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Validate that the PivotTable has at least one data field
            if (pivotTable.DataFields.Count == 0)
            {
                Console.WriteLine("Error: The PivotTable must contain at least one data field before calculations can be performed.");
                return;
            }

            // Refresh the PivotTable data
            pivotTable.RefreshData();

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
