// Title: Create an External VLOOKUP Formula with Correct Workbook Path in Aspose.Cells for .NET
// Description: Demonstrates how to generate an external workbook, build a VLOOKUP formula that references it using the proper "[FileName]Sheet!Range" syntax, set the main workbook's AbsolutePath, link the external source via CalculationOptions.LinkedDataSources, and evaluate the result with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# VLOOKUP external workbook | Excel external reference formula | AbsolutePath Aspose.Cells | LinkedDataSources calculation | .NET Excel lookup | external file path syntax
// Common Searches: Aspose.Cells VLOOKUP external file reference | How to set AbsolutePath for linked workbooks in C# | Calculate formulas with external data sources using Aspose.Cells | Excel VLOOKUP formula syntax for another workbook | C# example of external workbook lookup
// Developer Intent: Build and evaluate a VLOOKUP formula that pulls data from a separate Excel file using Aspose.Cells.
// Use Cases: Generate a main workbook that retrieves values from a lookup table stored in a different Excel file. | Ensure external references resolve correctly after saving both workbooks. | Reuse a single external workbook as a linked data source for multiple formulas.
// AI Prompts: Write C# code with Aspose.Cells to create a VLOOKUP formula that references an external workbook in the same folder, using the correct path syntax. | Explain how to configure AbsolutePath and CalculationOptions.LinkedDataSources so external VLOOKUP formulas are calculated automatically. | Provide a step‑by‑step tutorial for building an external lookup workbook, populating data, and linking it with a VLOOKUP formula in another workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVlookupExternal
{
    // Demonstrates how to generate an external workbook, build a VLOOKUP formula that references it using the proper "[FileName]Sheet!Range" syntax, set the main workbook's AbsolutePath, link the external source via CalculationOptions.LinkedDataSources, and evaluate the result with Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------------------
                // 1. Prepare the external workbook that will be referenced by VLOOKUP
                // -------------------------------------------------------------
                string externalFileName = "ExternalData.xlsx";
                string externalFullPath = Path.GetFullPath(externalFileName);

                // Ensure the directory exists (handle possible null from GetDirectoryName)
                string externalDir = Path.GetDirectoryName(externalFullPath);
                if (string.IsNullOrEmpty(externalDir))
                {
                    externalDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(externalDir))
                {
                    Directory.CreateDirectory(externalDir);
                }

                // Create a simple workbook with a table in Sheet1 (A1:B10)
                Workbook externalWb = new Workbook();
                Worksheet extSheet = externalWb.Worksheets[0];
                extSheet.Name = "Sheet1";

                // Populate lookup table: column A = keys, column B = values
                for (int i = 0; i < 10; i++)
                {
                    extSheet.Cells[i, 0].PutValue($"Key{i + 1}");
                    extSheet.Cells[i, 1].PutValue(i * 10); // some numeric value
                }

                // Save the external workbook to disk
                externalWb.Save(externalFullPath);

                // -------------------------------------------------------------
                // 2. Create the main workbook where the VLOOKUP formula will reside
                // -------------------------------------------------------------
                Workbook mainWb = new Workbook();
                Worksheet mainSheet = mainWb.Worksheets[0];
                mainSheet.Name = "Main";

                // Put a lookup key in A1 that exists in the external table
                mainSheet.Cells["A1"].PutValue("Key5");

                // -------------------------------------------------------------
                // 3. Set the AbsolutePath of the main workbook so that relative links work
                // -------------------------------------------------------------
                // AbsolutePath is used only for external links; set it to the folder of the external file
                mainWb.AbsolutePath = externalDir;

                // -------------------------------------------------------------
                // 4. Build the VLOOKUP formula string that references the external workbook
                // -------------------------------------------------------------
                // Syntax: =[ExternalFile.xlsx]SheetName!Range
                string externalFileOnly = Path.GetFileName(externalFullPath); // e.g., ExternalData.xlsx
                string vlookupFormula = $"=VLOOKUP(A1,'[{externalFileOnly}]Sheet1'!$A$1:$B$10,2,FALSE)";

                // -------------------------------------------------------------
                // 5. Assign the formula to a cell (C1)
                // -------------------------------------------------------------
                // Use the Formula property (compatible with all Aspose.Cells versions)
                mainSheet.Cells["C1"].Formula = vlookupFormula;

                // -------------------------------------------------------------
                // 6. Calculate the formula, providing the external workbook as a linked data source
                // -------------------------------------------------------------
                CalculationOptions calcOptions = new CalculationOptions
                {
                    LinkedDataSources = new Workbook[] { externalWb }
                };
                mainWb.CalculateFormula(calcOptions);

                // -------------------------------------------------------------
                // 7. Output the result and save the main workbook
                // -------------------------------------------------------------
                Console.WriteLine($"VLOOKUP result in C1: {mainSheet.Cells["C1"].Value}");

                string outputPath = Path.GetFullPath("MainWithVlookup.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                mainWb.Save(outputPath);
                Console.WriteLine($"Main workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
