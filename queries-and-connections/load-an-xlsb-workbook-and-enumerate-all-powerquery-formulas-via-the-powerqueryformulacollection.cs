// Title: C# – Load an XLSB workbook and list all Power Query formulas with Aspose.Cells
// Description: Sample code that opens an XLSB file (creating it if missing), accesses the workbook's DataMashup, iterates through the PowerQueryFormulaCollection to display each formula's name, definition and items, and saves the workbook to a new file.
// Keywords: Aspose.Cells XLSB | Power Query formula enumeration | DataMashup API | PowerQueryFormulaCollection C# | read Power Query mashup data | list Power Query items
// Common Searches: how to read Power Query formulas from an XLSB using Aspose.Cells | C# iterate PowerQueryFormulaCollection Aspose.Cells | check workbook for Power Query mashup data | save XLSB after enumerating Power Query formulas
// Developer Intent: Extract and display every Power Query formula stored in an XLSB workbook via Aspose.Cells.
// Use Cases: Verify that an incoming XLSB contains the expected Power Query definitions before processing. | Create an audit log of all Power Query formula names, definitions, and parameters. | Prepare the workbook for further automation by enumerating formulas, then editing or removing specific items.
// AI Prompts: Generate C# code that adds a new Power Query formula to an existing XLSB workbook using Aspose.Cells. | Write a method that extracts all Power Query formula definitions from a workbook and exports them to JSON. | Provide comprehensive error handling for missing DataMashup, empty PowerQueryFormulaCollection, or file‑access issues.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryDemo
{
    // Sample code that opens an XLSB file (creating it if missing), accesses the workbook's DataMashup, iterates through the PowerQueryFormulaCollection to display each formula's name, definition and items, and saves the workbook to a new file.
    class Program
    {
        static void Main()
        {
            const string sourcePath = "source.xlsb";
            const string outputPath = "source_modified.xlsb";

            try
            {
                // Ensure the source file exists; create an empty workbook if it does not.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"File '{sourcePath}' not found. Creating a new workbook.");
                    var newWb = new Workbook();
                    newWb.Save(sourcePath);
                }

                // Load the workbook (may contain Power Query formulas)
                Workbook workbook = new Workbook(sourcePath);

                // Access the mashup data of the workbook
                DataMashup mashup = workbook.DataMashup;

                // Verify that Power Query formulas are present
                if (mashup?.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
                {
                    Console.WriteLine("Power Query Formulas found:");

                    foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                    {
                        Console.WriteLine($"- Formula Name: {formula.Name}");
                        Console.WriteLine($"  Definition   : {formula.FormulaDefinition}");

                        // List items of the formula, if any
                        if (formula.PowerQueryFormulaItems != null && formula.PowerQueryFormulaItems.Count > 0)
                        {
                            Console.WriteLine("  Items:");
                            foreach (PowerQueryFormulaItem item in formula.PowerQueryFormulaItems)
                            {
                                Console.WriteLine($"    * {item.Name} = {item.Value}");
                            }
                        }

                        Console.WriteLine(); // blank line for readability
                    }
                }
                else
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                }

                // Save the workbook (optional, here we rewrite to a new file)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
