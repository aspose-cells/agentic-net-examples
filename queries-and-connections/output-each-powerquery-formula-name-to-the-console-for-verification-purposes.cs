// Title: List Power Query formula names from an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to open a .xlsx file, access its DataMashup, and print each PowerQueryFormula.Name to the console. | Extend the sample to also output the PowerQueryFormula.MExpression together with its name for every formula. | Add robust error handling that checks for missing files, absent DataMashup, and empty PowerQueryFormulas, logging clear messages for each scenario.
// Common Searches: aspocells c# list power query formulas in workbook | how to get Power Query M code from Excel using Aspose.Cells .NET | retrieve PowerQueryFormula names with Aspose.Cells DataMashup | C# read Power Query formulas from .xlsx file Aspose.Cells | enumerate Power Query formulas in Excel via Aspose.Cells API
// Tags: list PowerQueryFormula names Aspose.Cells | extract Power Query M expression DataMashup | C# read Power Query formulas from Excel workbook | handle missing DataMashup Aspose.Cells | save workbook after reading Power Query formulas

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    // // Loads an Excel file, accesses its DataMashup, iterates through the PowerQueryFormulas collection, writes each formula's Name to the console, and saves the workbook unchanged.
    public class ListPowerQueryFormulaNames
    {
        public static void Run()
        {
            try
            {
                string inputPath = "source.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook containing Power Query formulas
                Workbook workbook = new Workbook(inputPath);

                // Access the DataMashup object
                DataMashup mashup = workbook.DataMashup;

                // Verify that Power Query formulas exist
                if (mashup != null && mashup.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
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

                // Save the workbook (no modifications made, but required by lifecycle rules)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ListPowerQueryFormulaNames.Run();
        }
    }
}
