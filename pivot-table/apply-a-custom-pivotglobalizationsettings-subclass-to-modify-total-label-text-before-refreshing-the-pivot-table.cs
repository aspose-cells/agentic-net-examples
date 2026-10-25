// Title: How to subclass PivotGlobalizationSettings to change the total label text before refreshing a pivot table with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells custom PivotGlobalizationSettings to rename total label in pivot table | modify pivot table total label localization before RefreshData using Aspose.Cells .NET | apply custom globalization settings to pivot tables in Aspose.Cells C# example | change total label text in Excel pivot via Aspose.Cells PivotGlobalizationSettings subclass | refresh pivot data after customizing total label with Aspose.Cells .NET
// Tags: PivotGlobalizationSettings subclass Aspose.Cells | total label text customization pivot table C# | refresh pivot data after localization settings | Aspose.Cells pivot table localization customization | modify pivot total label before RefreshData

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // This example shows how to create a subclass of PivotGlobalizationSettings that overrides the default total label, assign the custom settings to a workbook, and then refresh and calculate the pivot table so the new label appears before saving the modified Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists before attempting to load it.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table on the worksheet.
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Retrieve the first pivot table.
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Refresh and calculate pivot data.
                try
                {
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while refreshing/calculating pivot table: {ex.Message}");
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
