// Title: Load an XLSX workbook with Aspose.Cells .NET and inspect its DataMashup Power Query formulas
// Description: Shows how to open an existing .xlsx file using Aspose.Cells Workbook, safely access the DataMashup property, handle null collections, enumerate Power Query formulas, display details of the first formula, and save the workbook to a new file.
// Keywords: Aspose.Cells | C# load workbook | Workbook DataMashup | Power Query formulas | DataMashup null check | Aspose.Cells .NET example | read Power Query mashup | save workbook Aspose.Cells | Excel DataMashup API | Aspose.Cells QueryTables
// Common Searches: Aspose.Cells how to get DataMashup from workbook | Read Power Query formulas with Aspose.Cells C# | Check if workbook.DataMashup is null | Enumerate PowerQueryFormulas collection Aspose.Cells | Save workbook after DataMashup inspection .NET | Get Power Query mashup data from Excel using Aspose
// Developer Intent: Open an existing Excel file, examine its DataMashup and Power Query formulas, and save the workbook after inspection.
// Use Cases: Verify whether a workbook contains Power Query mashup data before further processing. | Log names and item counts of all Power Query formulas for audit or documentation. | Extract details of the first Power Query formula and create a copy of the workbook with unchanged content.
// AI Prompts: Write C# code using Aspose.Cells to open an .xlsx file, check if workbook.DataMashup is null, and list all Power Query formula names. | Provide an example that loads a workbook, retrieves the first Power Query formula's details, and saves the workbook to a new file with Aspose.Cells .NET. | Explain best practices for handling null DataMashup or PowerQueryFormulas collections when working with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsDataMashupDemo
{
    // Shows how to open an existing .xlsx file using Aspose.Cells Workbook, safely access the DataMashup property, handle null collections, enumerate Power Query formulas, display details of the first formula, and save the workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the existing Excel file to be loaded
                string sourcePath = "input.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook using the string constructor (provided rule)
                Workbook workbook = new Workbook(sourcePath);

                // Access the DataMashup property (may be null if no mashup data)
                DataMashup dataMashup = workbook.DataMashup;
                if (dataMashup == null)
                {
                    Console.WriteLine("No DataMashup information found in the workbook.");
                }
                else
                {
                    // Retrieve Power Query formulas collection (may be null)
                    var powerQueryFormulas = dataMashup.PowerQueryFormulas;
                    if (powerQueryFormulas == null)
                    {
                        Console.WriteLine("PowerQueryFormulas collection is null.");
                    }
                    else
                    {
                        // Output basic information about the Power Query formulas
                        Console.WriteLine($"Number of Power Query formulas: {powerQueryFormulas.Count}");

                        // If there are any formulas, display details of the first one
                        if (powerQueryFormulas.Count > 0)
                        {
                            var firstFormula = powerQueryFormulas[0];
                            Console.WriteLine($"First Query Name: {firstFormula.Name}");
                            Console.WriteLine($"Number of items in first query: {firstFormula.PowerQueryFormulaItems.Count}");
                        }
                    }
                }

                // Save the workbook to a new file (using the provided Save method)
                string destPath = "output.xlsx";
                workbook.Save(destPath);
                Console.WriteLine($"Workbook saved to {destPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
