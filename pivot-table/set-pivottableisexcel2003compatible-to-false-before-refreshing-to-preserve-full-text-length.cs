// Title: Disable Excel 2003 compatibility for a PivotTable in Aspose.Cells for .NET to preserve full text length
// Common Searches: Aspose.Cells C# set PivotTable.IsExcel2003Compatible false before RefreshData | how to keep full text in Excel pivot table using Aspose.Cells .NET | disable Excel 2003 compatibility for pivot tables when processing workbooks with Aspose.Cells | refresh pivot table after changing IsExcel2003Compatible property in C# | Aspose.Cells example for modifying pivot table compatibility settings
// Tags: pivot table disable Excel2003 compatibility Aspose.Cells | set PivotTable.IsExcel2003Compatible false .NET | refresh pivot table after compatibility change Aspose.Cells | preserve full text length in Excel pivot tables C# | modify pivot table settings with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing XLSX workbook, checks for a pivot table on the first worksheet, sets the PivotTable.IsExcel2003Compatible property to false to retain the full text length, refreshes and recalculates the pivot data, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Disable Excel 2003 compatibility to keep full text length
            pivotTable.IsExcel2003Compatible = false;

            // Refresh and recalculate the pivot table after changing the setting
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
