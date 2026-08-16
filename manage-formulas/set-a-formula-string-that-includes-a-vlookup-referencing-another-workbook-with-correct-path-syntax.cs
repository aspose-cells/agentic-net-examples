// Title: Aspose.Cells C# – Set VLOOKUP Formula with Full Path to an External Workbook
// Description: Demonstrates how to create a secondary workbook, fill a lookup table, save it, and then assign a VLOOKUP formula in a primary workbook that references the external file using the required "'[full_path\File.xlsx]Sheet'!Range" syntax. The example also shows configuring CalculationOptions.LinkedDataSources so the formula is evaluated correctly.
// Keywords: Aspose.Cells external VLOOKUP | C# VLOOKUP across workbooks | full file path formula Aspose | LinkedDataSources calculation | reference another workbook in formula | Aspose.Cells formula property | temporary file path Excel lookup
// Common Searches: Aspose.Cells VLOOKUP external workbook path | C# set formula with full file path in Aspose.Cells | How to calculate VLOOKUP that points to another file | LinkedDataSources example Aspose.Cells | Excel VLOOKUP formula syntax for external file in .NET
// Developer Intent: Create and evaluate a VLOOKUP formula that pulls data from a separate workbook by specifying its absolute path.
// Use Cases: Generate a lookup workbook, populate key‑value pairs, and save it for reuse. | Insert a VLOOKUP formula in a cell of a main workbook that references the saved file using "'[path]Sheet'!Range" syntax. | Enable cross‑workbook calculation by adding the external workbook to CalculationOptions.LinkedDataSources.
// AI Prompts: Write C# code with Aspose.Cells that builds a VLOOKUP formula referencing an external workbook using a dynamic full file path and returns the calculated value. | Explain the role of CalculationOptions.LinkedDataSources when evaluating formulas that depend on another workbook in Aspose.Cells. | Provide a step‑by‑step tutorial for saving an external workbook, constructing the correct VLOOKUP string with path syntax, and executing the calculation in the main workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVlookupExternalDemo
{
    // Demonstrates how to create a secondary workbook, fill a lookup table, save it, and then assign a VLOOKUP formula in a primary workbook that references the external file using the required "'[full_path\File.xlsx]Sheet'!Range" syntax. The example also shows configuring CalculationOptions.LinkedDataSources so the formula is evaluated correctly.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create an external workbook that will be referenced by VLOOKUP
                // -----------------------------------------------------------------
                string externalPath = Path.Combine(Path.GetTempPath(), "ExternalData.xlsx");

                // Ensure the directory exists
                string externalDir = Path.GetDirectoryName(externalPath);
                if (!Directory.Exists(externalDir))
                {
                    Directory.CreateDirectory(externalDir);
                }

                Workbook externalWb = new Workbook();
                Worksheet extSheet = externalWb.Worksheets[0];
                extSheet.Name = "Data";

                // Populate a simple lookup table: Column A = keys, Column B = values
                extSheet.Cells["A1"].PutValue("Key");
                extSheet.Cells["B1"].PutValue("Value");
                extSheet.Cells["A2"].PutValue("Apple");
                extSheet.Cells["B2"].PutValue(10);
                extSheet.Cells["A3"].PutValue("Banana");
                extSheet.Cells["B3"].PutValue(20);
                extSheet.Cells["A4"].PutValue("Cherry");
                extSheet.Cells["B4"].PutValue(30);

                // Save the external workbook so it has a physical file path
                externalWb.Save(externalPath, SaveFormat.Xlsx);

                // ---------------------------------------------------------------
                // 2. Create the main workbook that will contain the VLOOKUP formula
                // ---------------------------------------------------------------
                Workbook mainWb = new Workbook();
                Worksheet mainSheet = mainWb.Worksheets[0];
                mainSheet.Name = "Main";

                // The lookup value we want to search for
                mainSheet.Cells["A2"].PutValue("Banana");

                // Build the VLOOKUP formula referencing the external workbook.
                // Syntax: =VLOOKUP(A2, '[full_path\ExternalData.xlsx]Data'!$A$2:$B$4, 2, FALSE)
                string vlookupFormula = $"=VLOOKUP(A2, '[{externalPath}]Data'!$A$2:$B$4, 2, FALSE)";

                // Set the formula (use the Formula property; Aspose.Cells will calculate it later)
                mainSheet.Cells["B2"].Formula = vlookupFormula;

                // ---------------------------------------------------------------
                // 3. Prepare calculation options so the external workbook can be used
                // ---------------------------------------------------------------
                CalculationOptions calcOptions = new CalculationOptions
                {
                    // Provide the external workbook as a linked data source.
                    LinkedDataSources = new Workbook[] { externalWb }
                };

                // Calculate the formula in the main workbook.
                mainWb.CalculateFormula(calcOptions);

                // ---------------------------------------------------------------
                // 4. Output the result to console and save the main workbook
                // ---------------------------------------------------------------
                Console.WriteLine("VLOOKUP result (should be 20): " + mainSheet.Cells["B2"].Value);

                string mainPath = Path.Combine(Path.GetTempPath(), "MainWithVlookup.xlsx");
                mainWb.Save(mainPath, SaveFormat.Xlsx);
                Console.WriteLine("Main workbook saved to: " + mainPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
