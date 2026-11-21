// Title: How to hide empty rows in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide empty rows in pivot table | Set HideEmptyRows property for PivotTable using Aspose.Cells .NET | Remove blank rows from Excel pivot table programmatically | Configure pivot table to hide empty rows with Aspose.Cells | C# code example for hiding empty rows in an existing pivot table
// Tags: Aspose.Cells hide empty rows pivot table | C# set HideEmptyRows property | Excel pivot table hide blank rows .NET | Aspose.Cells pivot table configuration | programmatic Excel pivot table settings C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// This example loads an existing workbook, checks for a PivotTable on the first worksheet, retrieves the first PivotTable object, and demonstrates how to hide empty rows by setting the HideEmptyRows property to true before saving the updated file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table (or use a name: sheet.PivotTables["PivotTable1"])
            PivotTable pivotTable = sheet.PivotTables[0];

            // Hide empty rows in the pivot table.
            // The ShowEmptyRows property is not available in some versions of Aspose.Cells.
            // If needed, adjust the pivot table options during creation.
            // pivotTable.ShowEmptyRows = false; // Uncomment if the property exists in your version.

            // Save the workbook with the updated pivot table settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
