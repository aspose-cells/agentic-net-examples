// Title: Clear all row fields from a PivotTable in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove row fields from PivotTable | reset pivot table layout by clearing row fields in .NET | how to clear PivotTable row fields programmatically with Aspose.Cells | C# example to clear all row fields of a PivotTable | modify existing Excel workbook pivot row fields using Aspose
// Tags: Aspose.Cells PivotTable RowFields.Clear method | clear pivot row fields C# | reset pivot table layout Aspose.Cells | modify pivot table row fields .NET | load workbook edit pivot Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook, accesses the first worksheet's first PivotTable, clears all row fields with the RowFields.Clear method to reset the PivotTable layout, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password is supplied; adjust LoadOptions if a password is required)
            var loadOptions = new LoadOptions(); // add password here if needed: loadOptions.Password = "yourPassword";
            var workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable on the worksheet
            if (worksheet.PivotTables.Count > 0)
            {
                var pivotTable = worksheet.PivotTables[0];

                // Remove all row fields to reset the PivotTable layout
                pivotTable.RowFields.Clear();
            }
            else
            {
                Console.WriteLine("Warning: No PivotTables found on the first worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
