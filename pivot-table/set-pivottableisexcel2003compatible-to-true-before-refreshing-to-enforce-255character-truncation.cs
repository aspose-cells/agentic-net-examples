// Title: Set PivotTable.IsExcel2003Compatible to true and refresh the pivot table with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set PivotTable.IsExcel2003Compatible property in C# | Refresh a pivot table after enabling Excel 2003 compatibility with Aspose.Cells | Truncate pivot table field values to 255 characters using Aspose.Cells .NET | Load a workbook, change pivot table compatibility mode, and save with Aspose.Cells | Enable Excel 2003 mode for pivot tables before calling RefreshData in C#
// Tags: Excel2003 compatibility flag for pivot tables | pivot table refresh using Aspose.Cells .NET | enforce 255‑character truncation in pivot tables | save workbook after pivot table update | modify pivot table compatibility mode Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, accesses the first worksheet's first pivot table, sets its IsExcel2003Compatible property to true to enforce 255‑character truncation, refreshes the pivot table, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet that contains the pivot table (first worksheet by default)
            Worksheet sheet = workbook.Worksheets[0]; // adjust index as needed

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table in the sheet
            PivotTable pivotTable = sheet.PivotTables[0]; // adjust index if multiple pivots

            // Enable Excel 2003 compatibility to enforce 255‑character truncation
            pivotTable.IsExcel2003Compatible = true;

            // Refresh the pivot table so the setting takes effect
            pivotTable.RefreshData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
