// Title: Set row grand totals visible and hide column grand totals in an Excel pivot table with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to display row grand totals in a pivot table | C# code to hide column grand totals in Excel pivot using Aspose.Cells | Enable row grand totals and disable column grand totals with Aspose.Cells .NET | Modify pivot table total settings programmatically in Aspose.Cells | Show row totals only in Excel pivot table via Aspose.Cells API
// Tags: Aspose.Cells pivot row totals display | Aspose.Cells pivot column totals hide | C# configure pivot table totals Aspose.Cells | Excel pivot totals configuration .NET API | programmatic pivot table total control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads an existing workbook, checks for a worksheet and a pivot table, retrieves the first pivot table, enables row grand totals, disables column grand totals, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Show row grand totals
            pivotTable.ShowRowGrandTotals = true;

            // Hide column grand totals
            pivotTable.ShowColumnGrandTotals = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
