// Title: Load an XLSX workbook and read Power Query formulas via Workbook.DataMashup using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a .xlsx file with Aspose.Cells, accesses the Workbook.DataMashup property, and iterates over all PowerQueryFormulas. | Show how to verify that DataMashup is not null, output the total count of Power Query formulas, and print the name and item count of the first formula. | Provide an example that saves the workbook after inspecting its DataMashup information, including handling a missing source file.
// Common Searches: Aspose.Cells C# load workbook and get Power Query formulas count | how to check for DataMashup existence in an Excel file using Aspose.Cells | retrieve first PowerQueryFormula name from workbook with Aspose.Cells .NET | save workbook after reading DataMashup property in Aspose.Cells
// Tags: load xlsx workbook Aspose.Cells | access Workbook.DataMashup API | enumerate PowerQueryFormulas C# | null DataMashup check Aspose.Cells | save workbook after DataMashup inspection

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsDataMashupDemo
{
    // The example demonstrates loading an existing XLSX file into an Aspose.Cells Workbook, safely accessing its DataMashup property, enumerating Power Query formulas, printing details of the first formula, and saving the workbook to a new file while handling missing files and absent DataMashup data.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the existing Excel file that will be loaded.
                string sourcePath = "input.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook from the file.
                Workbook workbook = new Workbook(sourcePath);

                // Access the DataMashup property safely.
                DataMashup dataMashup = workbook.DataMashup;
                if (dataMashup == null)
                {
                    Console.WriteLine("The workbook does not contain any DataMashup information.");
                }
                else
                {
                    // Retrieve the collection of Power Query formulas.
                    var powerQueryFormulas = dataMashup.PowerQueryFormulas;

                    // Output basic information about the Power Query formulas.
                    Console.WriteLine($"Number of Power Query formulas: {powerQueryFormulas?.Count ?? 0}");

                    // If there are any Power Query formulas, display details of the first one.
                    if (powerQueryFormulas != null && powerQueryFormulas.Count > 0)
                    {
                        var firstFormula = powerQueryFormulas[0];
                        Console.WriteLine($"First Query Name: {firstFormula.Name}");
                        Console.WriteLine($"Number of items in first query: {firstFormula.PowerQueryFormulaItems?.Count ?? 0}");
                    }
                }

                // Optionally, save the workbook to a new file.
                string destPath = "output.xlsx";
                workbook.Save(destPath);
                Console.WriteLine("Workbook loaded, DataMashup accessed, and file saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
