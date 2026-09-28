// Title: Disable column grand totals in an Excel PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an existing workbook, locates the first PivotTable, and sets ShowColumnGrandTotals to false. | Show how to add error handling for a missing input file and for worksheets without PivotTables while turning off column grand totals. | Demonstrate saving the workbook to a new file after programmatically disabling column grand totals in a PivotTable.
// Common Searches: Aspose.Cells C# hide column grand totals in pivot table | set ShowColumnGrandTotals false programmatically .NET | remove column grand total from Excel pivot using Aspose.Cells | how to turn off column grand totals in a PivotTable with C# | Aspose.Cells pivot table display options column totals
// Tags: Aspose.Cells ShowColumnGrandTotals property | C# disable pivot table column grand totals | Aspose.Cells modify pivot table display settings | Excel pivot column totals Aspose.Cells .NET | programmatic pivot table formatting C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, accesses the first worksheet and its first PivotTable, sets the ShowColumnGrandTotals property to false to hide column grand totals, and saves the result to a new file. It also includes checks for a missing input file and for worksheets that contain no PivotTables.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found on the first worksheet.");
                return;
            }

            // Access the first PivotTable on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Disable column grand totals
            pivotTable.ShowColumnGrandTotals = false;

            // Save the workbook with the updated PivotTable settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
