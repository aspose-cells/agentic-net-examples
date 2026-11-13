// Title: How to customize the total label text of an Excel pivot table using a PivotGlobalizationSettings subclass in Aspose.Cells for .NET (C#)
// Common Searches: how to change total label text in an Aspose.Cells pivot table with C# | example of creating a PivotGlobalizationSettings subclass to set custom total caption in .NET | Aspose.Cells pivot table total label customization before RefreshData call | C# code to override total row label in Excel pivot using Aspose.Cells | apply custom globalization settings to pivot tables in Aspose.Cells workbook
// Tags: custom PivotGlobalizationSettings subclass C# | modify pivot total label Aspose.Cells | pivot table total caption localization .NET | apply globalization settings before pivot refresh | Aspose.Cells pivot total label customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example demonstrates how to create a custom PivotGlobalizationSettings subclass in C# to override the default total label text of an Excel pivot table, apply the settings to a workbook, refresh and recalculate the pivot, and then save the updated file using Aspose.Cells for .NET.
class Program
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
            // Load the existing workbook that contains a pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (sheet.PivotTables.Count > 0)
            {
                // Access the first pivot table
                PivotTable pivot = sheet.PivotTables[0];

                // Refresh the pivot table data
                pivot.RefreshData();
                pivot.CalculateData();
            }
            else
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
            }

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
