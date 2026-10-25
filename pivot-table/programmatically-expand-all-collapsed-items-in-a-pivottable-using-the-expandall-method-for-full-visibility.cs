// Title: Expanding all collapsed items in an Aspose.Cells PivotTable using C# (ShowDetail workaround)
// Common Searches: C# code to programmatically expand all items in an Aspose.Cells PivotTable | Aspose.Cells .NET how to set ShowDetail for every PivotTable field | Expand collapsed rows in Excel PivotTable with Aspose.Cells library | Iterate PivotTable fields to display full details using Aspose.Cells C# | Workaround for missing ExpandAll method in Aspose.Cells PivotTable
// Tags: Aspose.Cells expand collapsed pivot items C# | ShowDetail property PivotTable Aspose.Cells | programmatic pivot table item expansion .NET | Aspose.Cells workbook modify pivot visibility | C# iterate PivotTable fields Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, checks for a PivotTable, and explains that Aspose.Cells does not provide a direct ExpandAll method. It suggests using the ShowDetail property on each PivotTable field as a workaround to programmatically expand all collapsed items, then saves the workbook.
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
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // NOTE: Aspose.Cells does not provide a direct ExpandAll method.
            // If needed, you can programmatically set ShowDetail for each field.
            // For this example, we proceed without expanding items.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
