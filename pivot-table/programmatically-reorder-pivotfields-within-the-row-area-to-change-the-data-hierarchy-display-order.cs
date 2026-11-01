// Title: Programmatically move row PivotFields to change hierarchy order in an Excel pivot table with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to change the order of row fields in an Excel pivot table | how to move a pivot table row field to a different position using Aspose.Cells .NET | reorder row hierarchy in Excel pivot table programmatically with Aspose.Cells | Aspose.Cells move PivotFields in row area after loading password‑protected workbook | sample code for shifting pivot table row fields using Aspose.Cells in C#
// Tags: aspocells pivotrowfields move c# | excel pivot table row hierarchy reorder .net | load password protected workbook aspocells | pivot table refresh calculate aspocells | c# reorder pivotfields programmatically | aspocells row area field sequence

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (handling optional password protection), accesses the first worksheet, verifies a pivot table exists, and if there are at least three row fields, moves the first row field to the third position using PivotFields.Move. It then refreshes and recalculates the pivot table before saving the modified workbook to a new file.
class Program
{
    static void Main(string[] args)
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Try to load workbook without password
            workbook = new Workbook(inputPath);
        }
        catch (CellsException)
        {
            // Assume password is required and prompt the user
            Console.WriteLine("The workbook is password protected.");
            Console.Write("Enter password: ");
            string password = Console.ReadLine() ?? string.Empty;

            var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            try
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook with password: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access first worksheet
            var worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            var pivotTable = worksheet.PivotTables[0];

            // Reorder row fields if enough fields are present
            if (pivotTable.RowFields.Count > 2)
            {
                pivotTable.RowFields.Move(0, 2);
            }
            else
            {
                Console.WriteLine("Insufficient row fields to change order.");
            }

            // Refresh and recalculate pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing pivot table: {ex.Message}");
        }
    }
}
