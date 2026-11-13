// Title: Disable Excel 2003 compatibility for a PivotTable and refresh it using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PivotTable.IsExcel2003Compatible to false before RefreshData | How to keep full text length in a PivotTable when exporting with Aspose.Cells | Refresh and recalculate a PivotTable after changing compatibility mode in .NET | Disable Excel 2003 compatibility for PivotTables in a workbook with Aspose.Cells | C# code to turn off Excel 2003 compatibility for a PivotTable and save workbook
// Tags: pivot table disable Excel2003 compatibility Aspose.Cells | refresh pivot table after property change .NET | preserve full text length in pivot tables Aspose | set IsExcel2003Compatible false C# | calculate pivot data after compatibility mode change

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, accesses the first worksheet's first pivot table, disables Excel 2003 compatibility to retain full text length, refreshes and recalculates the pivot data, and saves the modified workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Disable Excel 2003 compatibility to preserve full text length
            pivotTable.IsExcel2003Compatible = false;

            // Refresh and recalculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
