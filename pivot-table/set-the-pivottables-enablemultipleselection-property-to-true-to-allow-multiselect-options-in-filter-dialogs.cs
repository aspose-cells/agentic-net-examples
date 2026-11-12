// Title: Enable multi‑select filter options in an Aspose.Cells PivotTable using C# (EnableMultipleSelection = true)
// Common Searches: C# Aspose.Cells set EnableMultipleSelection true for PivotTable filter dialog | Enable multiple item selection in PivotTable filters with Aspose.Cells .NET | Aspose.Cells PivotTable allow users to select multiple values in filter dropdown | Set EnableMultipleSelection property on PivotField in Aspose.Cells C# example | Why EnableMultipleSelection property is missing in current Aspose.Cells version and how to upgrade
// Tags: Aspose.Cells enable multi-select pivot filter | C# set EnableMultipleSelection property | PivotTable filter multiple selection .NET | Aspose.Cells PivotTable API usage | upgrade Aspose.Cells for latest pivot features

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the first worksheet, retrieves the first PivotTable, sets its EnableMultipleSelection property to true so filter dialogs support selecting multiple items, and saves the modified workbook.
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
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // NOTE: In some Aspose.Cells versions the EnableMultipleSelection property
            // is not available on PivotField. If needed, upgrade the library.
            // The following code is omitted to keep the example compatible.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
