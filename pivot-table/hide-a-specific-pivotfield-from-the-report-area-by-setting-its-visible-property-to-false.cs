// Title: Hide a specific row field in an Excel pivot table with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells to hide the "Category" row field in an existing pivot table and save the workbook. | Show how to programmatically remove a pivot row field from the report area with Aspose.Cells for .NET. | Create a C# example that verifies a workbook contains a pivot table, finds a row field by name, hides it, and writes the updated file.
// Common Searches: aspnet cells hide pivot row field Category C# | remove specific field from Excel pivot table using Aspose.Cells .NET | how to programmatically hide a pivot field in an existing workbook with Aspose.Cells | C# Aspose.Cells example to delete a row field from a pivot table | Excel pivot table field visibility Aspose.Cells API
// Tags: aspnet-cells hide pivot row field | aspnet-cells remove pivot field | csharp aspnet-cells modify pivot table layout | excel pivot table field visibility aspnet-cells | aspnet-cells pivot table rowfield delete

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, accesses the first worksheet's first pivot table, locates the row field named "Category", removes it from the RowFields collection to hide it from the report area, and saves the modified workbook as output.xlsx, with file existence checks and exception handling.
class HidePivotFieldExample
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Assume the first pivot table on the sheet is the target
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Retrieve the row field named "Category"
            PivotField pivotField = pivotTable.RowFields["Category"];
            if (pivotField == null)
            {
                Console.WriteLine("Pivot field \"Category\" not found in row fields.");
                return;
            }

            // Hide the field by removing it from the row fields collection
            pivotTable.RowFields.Remove(pivotField);

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
