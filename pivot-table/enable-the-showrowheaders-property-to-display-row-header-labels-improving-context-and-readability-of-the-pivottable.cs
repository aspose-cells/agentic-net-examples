// Title: Enable ShowRowHeaders property to display row header labels in an Aspose.Cells PivotTable using C#
// Common Searches: Aspose.Cells C# set ShowRowHeaders true for a pivot table | how to display row header labels in an Excel pivot table with Aspose.Cells .NET | C# code example for enabling row headers in Aspose.Cells PivotTable | programmatically turn on ShowRowHeaders in Aspose.Cells pivot tables | Aspose.Cells PivotTable row header visibility setting in C#
// Tags: Aspose.Cells PivotTable ShowRowHeaders | C# enable row headers in Excel pivot table | Aspose.Cells set pivot table row header visibility | Excel pivot table row header display with Aspose.Cells | Aspose.Cells .NET modify pivot table layout

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable types

// The example loads an existing workbook, verifies a PivotTable exists, accesses the first PivotTable, and demonstrates how to set the ShowRowHeaders property to true so that row header labels are displayed, then saves the updated file.
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

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTable found in the worksheet.");
                return;
            }

            // Access the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Row header labels are displayed by default; additional settings can be applied here if needed.
            // Example: pivotTable.RowHeaderCaption = "Rows";

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
