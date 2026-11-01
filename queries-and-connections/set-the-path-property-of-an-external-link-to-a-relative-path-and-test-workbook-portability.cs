// Title: How to set a relative path for an external workbook link and test portability with Aspose.Cells for .NET
// AI Prompts: Create a source workbook, fill cells A1:A5 with numbers 1‑5, save it in a subfolder, then build a target workbook that inserts a SUM formula referencing the source workbook using a relative path (Data/Source.xlsx) and save the target file. | Reload the target workbook, read the formula in B2 to confirm the relative path is retained, and evaluate the formula to verify the external link works after reload.
// Common Searches: Aspose.Cells set external link to relative file path in C# | verify external workbook reference remains after saving with Aspose.Cells | C# example for using relative paths in Excel formulas with Aspose.Cells | how to calculate a formula that references another workbook using a relative path in Aspose.Cells | make Excel workbook portable by using relative external links in Aspose.Cells .NET
// Tags: Aspose.Cells external link relative path | C# set workbook external reference | Aspose.Cells calculate formula with external workbook | Excel workbook portability Aspose.Cells | relative file path in Excel formula C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDemo
{
    // The sample creates a source workbook with values 1‑5, saves it under a Data folder, then creates a target workbook that uses a SUM formula referencing the source workbook via a relative path (Data/Source.xlsx). After saving, the target workbook is reloaded to confirm the formula retains the relative path and evaluates correctly, demonstrating workbook portability.
    class Program
    {
        static void Main()
        {
            try
            {
                // Define paths
                string baseFolder = Path.Combine(Directory.GetCurrentDirectory(), "WorkbookDemo");
                string sourceFolder = Path.Combine(baseFolder, "Data");
                string sourceFile = Path.Combine(sourceFolder, "Source.xlsx");
                string targetFile = Path.Combine(baseFolder, "Target.xlsx");

                // Ensure directories exist
                Directory.CreateDirectory(sourceFolder);
                Directory.CreateDirectory(baseFolder);

                // -----------------------------------------------------------------
                // Step 1: Create a source workbook that will be referenced externally
                // -----------------------------------------------------------------
                Workbook sourceWb = new Workbook();
                Worksheet sourceWs = sourceWb.Worksheets[0];
                sourceWs.Name = "Sheet1";

                // Populate some data in the source workbook (A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    sourceWs.Cells[i, 0].PutValue(i + 1); // Values 1..5
                }

                // Save the source workbook
                sourceWb.Save(sourceFile);

                // -----------------------------------------------------------------
                // Step 2: Create the target workbook that will contain an external link
                // -----------------------------------------------------------------
                Workbook targetWb = new Workbook();
                Worksheet targetWs = targetWb.Worksheets[0];
                targetWs.Name = "Main";

                // Build a relative path to the source workbook
                string relativePath = Path.Combine("Data", "Source.xlsx").Replace("\\", "/");

                // Insert a formula that references the source workbook using the relative path
                targetWs.Cells["B2"].Formula = $"=SUM('[" + relativePath + "]Sheet1'!A1:A5)";

                // -----------------------------------------------------------------
                // Step 3: Save the target workbook
                // -----------------------------------------------------------------
                targetWb.Save(targetFile);

                // -----------------------------------------------------------------
                // Step 4: Load the saved workbook to verify that the relative path is retained
                // -----------------------------------------------------------------
                if (!File.Exists(targetFile))
                    throw new FileNotFoundException("Target workbook not found.", targetFile);

                Workbook loadedWb = new Workbook(targetFile);
                Worksheet loadedWs = loadedWb.Worksheets[0];

                // Output the formula to confirm the relative path
                Console.WriteLine("Formula in B2 after reload: " + loadedWs.Cells["B2"].Formula);

                // Evaluate the formula (requires the source file to be accessible)
                loadedWb.CalculateFormula();
                Console.WriteLine("Calculated value in B2: " + loadedWs.Cells["B2"].Value);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
