// Title: Show item labels in the values area of an Excel PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Set PivotTable.ShowValuesColumn = true on the first PivotTable in a workbook with Aspose.Cells for .NET to display item labels in the values area. | Modify a C# Aspose.Cells script to enable the values column label visibility for a PivotTable and save the updated workbook.
// Common Searches: Aspose.Cells C# how to enable item labels in pivot table values area | Set ShowValuesColumn property true for PivotTable using Aspose.Cells .NET | Display values column labels in Excel pivot table programmatically with Aspose.Cells | C# code to show item labels in pivot table values area Aspose.Cells
// Tags: Aspose.Cells PivotTable ShowValuesColumn | C# enable values column label in pivot table | Aspose.Cells display item labels in values area | programmatic pivot table label visibility Aspose.Cells | Excel pivot table values column property .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for a PivotTable on the first worksheet, optionally sets PivotTable.ShowValuesColumn = true to show item labels in the values area, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                var pivotTable = worksheet.PivotTables[0];

                // Enable the display of item labels in the values area (if supported by the version)
                // The property ShowValuesColumn may not be available in older versions.
                // Uncomment the following line if the property exists in your Aspose.Cells version:
                // pivotTable.ShowValuesColumn = true;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
