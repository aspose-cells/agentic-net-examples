// Title: Configure an Aspose.Cells PivotTable in C# to show row grand totals while hiding column grand totals
// Common Searches: Aspose.Cells C# set pivot table RowGrand true ColumnGrand false | How to hide column grand totals in an Excel pivot table using Aspose.Cells for .NET | Show only row totals in a pivot table with Aspose.Cells API | C# code to modify pivot table grand total visibility in an existing workbook | Aspose.Cells pivot table grand total settings example
// Tags: Aspose.Cells pivot table grand total settings | C# set RowGrand property Aspose.Cells | C# hide ColumnGrand Aspose.Cells | modify pivot table totals .NET | Excel workbook pivot table manipulation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, accesses the first worksheet's first pivot table, enables row grand totals, disables column grand totals, and saves the updated workbook.
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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Show row grand totals
            pivotTable.RowGrand = true;

            // Hide column grand totals
            pivotTable.ColumnGrand = false;

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
