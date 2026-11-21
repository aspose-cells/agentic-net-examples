// Title: How to delete a specific PivotTable from an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove a pivot table by object reference | Delete a particular PivotTable from a worksheet with Aspose.Cells .NET | C# code to remove the first PivotTable in an Excel file using Aspose.Cells | Worksheet.PivotTables.Remove method example in Aspose.Cells | Programmatically delete a pivot table in an .xlsx file using Aspose.Cells
// Tags: Aspose.Cells PivotTables.Remove method | C# delete pivot table from .xlsx | remove specific PivotTable object Aspose.Cells | Excel worksheet pivot table deletion .NET | programmatic pivot table removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for a worksheet named "Sheet1" and ensures it contains at least one pivot table, then retrieves the first PivotTable object and removes it with worksheet.PivotTables.Remove before saving the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the target worksheet (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"];
            if (worksheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Sheet1\" not found.");
                return;
            }

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table and remove it
            PivotTable pivotTable = worksheet.PivotTables[0];
            worksheet.PivotTables.Remove(pivotTable);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table removed successfully. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
