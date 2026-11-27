// Title: Changing the data source range of a pivot table in an existing Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change pivot table source range in existing workbook | Update pivot table data source to new range programmatically with Aspose.Cells | Refresh pivot table after modifying data source using Aspose.Cells .NET | Load Excel file and adjust pivot table source range with Aspose.Cells C# | Aspose.Cells ChangeDataSource method example for pivot tables
// Tags: Aspose.Cells ChangeDataSource method for pivot tables | update pivot table source range in Excel using Aspose.Cells | refresh pivot table after data source modification Aspose.Cells | load existing workbook and modify pivot table C# | adjust Excel pivot table data source programmatically .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, accesses its first worksheet, verifies that a pivot table exists, selects the first pivot table, changes its data source to the range "A1:D100" using the ChangeDataSource method, refreshes the pivot table, and saves the updated workbook to a new file, with basic error handling.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Locate the pivot table (by index; you can also use the name)
            PivotTable pivotTable = worksheet.PivotTables[0];
            // Example using name:
            // PivotTable pivotTable = worksheet.PivotTables["MyPivotTable"];

            // Define the new data source range as a string (e.g., A1:D100)
            string newDataSource = "A1:D100";

            // Change the pivot table's data source to the new range (expects string[])
            pivotTable.ChangeDataSource(new[] { newDataSource });

            // Refresh the pivot table to apply the new source
            pivotTable.RefreshData();

            // Save the workbook with the updated pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
