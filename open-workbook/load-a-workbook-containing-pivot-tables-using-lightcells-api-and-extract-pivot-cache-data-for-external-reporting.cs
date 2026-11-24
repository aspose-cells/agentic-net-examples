// Title: Load an Excel workbook and enumerate all pivot tables across worksheets using Aspose.Cells for .NET
// Common Searches: how to enumerate pivot tables in an Excel file using Aspose.Cells C# | Aspose.Cells .NET list all pivot tables in a workbook | C# code to read pivot table names from each worksheet with Aspose.Cells | retrieve pivot table information from an Excel workbook using Aspose.Cells library | iterate through worksheets and get pivot table details Aspose.Cells C#
// Tags: enumerate pivot tables Aspose.Cells .NET | read pivot table metadata Excel C# | iterate worksheets Aspose.Cells workbook | extract pivot table names from Excel workbook | load workbook containing pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example verifies the existence of an input Excel file, loads it with Aspose.Cells, walks through every worksheet, and prints each pivot table's name along with its sheet, demonstrating how to enumerate pivot tables in a workbook using C#.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook that may contain pivot tables.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the current worksheet.
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Output basic information about the pivot table.
                    Console.WriteLine($"Pivot Table: {pivotTable.Name} (Worksheet: {sheet.Name})");
                    Console.WriteLine(new string('-', 50));
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
