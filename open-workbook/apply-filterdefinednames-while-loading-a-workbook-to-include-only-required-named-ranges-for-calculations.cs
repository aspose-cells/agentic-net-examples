// Title: Load an Excel workbook with Aspose.Cells for .NET while retaining only selected named ranges
// AI Prompts: Load a workbook, iterate through Worksheets.Names, and delete any defined name that is not present in a supplied HashSet using Aspose.Cells. | After pruning unwanted named ranges, call workbook.CalculateFormula() and save the modified workbook to a new file. | Add error handling to verify the source Excel file exists before filtering named ranges with Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to load a workbook and remove all defined names except a few specific ones | C# filter named ranges in an Excel file using Aspose.Cells before recalculating formulas | Remove unwanted Excel defined names during workbook load with Aspose.Cells and save the result
// Tags: selective named range retention Aspose.Cells .NET | remove extraneous defined names C# | recalculate workbook formulas after name changes Aspose.Cells | load Excel workbook with filtered named ranges Aspose.Cells | hashset based named range filtering .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads input.xlsx with Aspose.Cells, removes every defined name except "SalesData" and "TaxRate", recalculates all formulas, and saves the cleaned workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Define the named ranges that need to be retained
                var requiredNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "SalesData",
                    "TaxRate"
                };

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Remove defined names that are not required
                var definedNames = workbook.Worksheets.Names; // Collection of defined names
                for (int i = definedNames.Count - 1; i >= 0; i--)
                {
                    var name = definedNames[i];
                    // The Name class exposes the name string via the Text property
                    if (!requiredNames.Contains(name.Text))
                    {
                        definedNames.RemoveAt(i);
                    }
                }

                // Recalculate all formulas after modifications
                workbook.CalculateFormula();

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
