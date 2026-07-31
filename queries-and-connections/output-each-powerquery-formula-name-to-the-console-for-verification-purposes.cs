// Title: C# – List Power Query Formula Names in an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file using Aspose.Cells, accesses the workbook's DataMashup, enumerates the PowerQueryFormulas collection, prints each formula's Name to the console, and optionally saves the unchanged workbook.
// Keywords: Aspose.Cells Power Query formulas | C# list PowerQueryFormula names | DataMashup PowerQueryFormulas enumeration | read Power Query names from Excel | Aspose.Cells console output example
// Common Searches: how to list Power Query formulas in Excel using Aspose.Cells C# | enumerate PowerQueryFormula objects with Aspose.Cells | C# code to get Power Query formula names from a workbook | Aspose.Cells DataMashup PowerQueryFormulas sample
// Developer Intent: Retrieve and display every Power Query formula name stored in an Excel workbook.
// Use Cases: Audit a workbook to confirm required Power Query queries are present. | Generate documentation that lists all Power Query formulas used in a file. | Detect unexpected or leftover Power Query formulas after automated workbook generation.
// AI Prompts: Create C# code with Aspose.Cells that prints all Power Query formula names and handles an empty collection gracefully. | Explain why saving the workbook after listing formulas is optional and show the save call. | Show how to filter Power Query formulas by a naming pattern (e.g., prefix "Sales_") using the DataMashup object.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    // Loads an Excel file using Aspose.Cells, accesses the workbook's DataMashup, enumerates the PowerQueryFormulas collection, prints each formula's Name to the console, and optionally saves the unchanged workbook.
    public class ListPowerQueryFormulaNames
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that may contain Power Query formulas
                Workbook workbook = new Workbook(inputPath);

                // Access the DataMashup object
                DataMashup mashup = workbook.DataMashup;

                // Check if any Power Query formulas exist
                if (mashup?.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
                {
                    Console.WriteLine("Power Query Formula Names:");
                    foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                    {
                        // Output each formula's name
                        Console.WriteLine(formula.Name);
                    }
                }
                else
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                }

                // Save the workbook (optional, as we didn't modify it)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
