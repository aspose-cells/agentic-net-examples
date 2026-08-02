// Title: Aspose.Cells for .NET – Create an External Named Range that References Another Workbook
// Description: C# example that builds a source workbook, adds an external link and external name, defines a named range in a second workbook using SetRefersTo, applies the range in a SUM formula, evaluates the formula with CalculationOptions.LinkedDataSources, and saves both files. Demonstrates cross‑workbook formulas with Aspose.Cells.
// Keywords: Aspose.Cells external named range | cross workbook reference .NET | C# Aspose.Cells external link | SetRefersTo external range | CalculationOptions LinkedDataSources | Excel formula across files | SUM external named range | Aspose.Cells tutorial
// Common Searches: how to reference a range in another Excel file using Aspose.Cells | Aspose.Cells create external named range C# | calculate formulas that use external workbook data Aspose.Cells | add external link and external name Aspose.Cells .NET | use SetRefersTo for cross‑file named range
// Developer Intent: Define a named range in a workbook that points to a range in a different workbook and use it in formulas.
// Use Cases: Consolidate shared data by linking a reporting workbook to a master data workbook via an external named range. | Build a financial model that sums values from a separate data file without hard‑coding the external reference. | Perform VLOOKUP or INDEX/MATCH lookups against a centralized lookup table stored in another workbook.
// AI Prompts: Generate C# code with Aspose.Cells to create an external named range that points to A1:A5 in "Data.xlsx" and use it in a VLOOKUP formula. | Explain how to set CalculationOptions.LinkedDataSources so formulas referencing external named ranges are evaluated correctly. | Show how to add multiple external links and external names in one workbook and reference each through separate named ranges.

using System;
using Aspose.Cells;

namespace AsposeCellsExternalNamedRangeDemo
{
    // C# example that builds a source workbook, adds an external link and external name, defines a named range in a second workbook using SetRefersTo, applies the range in a SUM formula, evaluates the formula with CalculationOptions.LinkedDataSources, and saves both files. Demonstrates cross‑workbook formulas with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create an external workbook that will hold the source data
            // -----------------------------------------------------------------
            string externalFileName = "ExternalData.xlsx";
            Workbook externalWb = new Workbook();
            Worksheet externalSheet = externalWb.Worksheets[0];
            externalSheet.Name = "Sheet1";

            // Populate some sample values
            externalSheet.Cells["A1"].PutValue(10);
            externalSheet.Cells["A2"].PutValue(20);
            externalSheet.Cells["A3"].PutValue(30);

            // Save the external workbook (lifecycle rule: use provided save)
            externalWb.Save(externalFileName);

            // ---------------------------------------------------------------
            // Step 2: Create the main workbook where the named range will be used
            // ---------------------------------------------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainSheet = mainWb.Worksheets[0];
            mainSheet.Name = "MainSheet";

            // ---------------------------------------------------------------
            // Step 3: Add an external link to the external workbook
            // ---------------------------------------------------------------
            string[] externalSheetNames = new string[] { "Sheet1" };
            int externalLinkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFileName, externalSheetNames);
            ExternalLink externalLink = mainWb.Worksheets.ExternalLinks[externalLinkIndex];

            // Add an external name that points to the range we want to reference
            externalLink.AddExternalName("ExtRange", "=Sheet1!$A$1:$A$3");

            // ---------------------------------------------------------------
            // Step 4: Create a named range in the main workbook that references the external range
            // ---------------------------------------------------------------
            int nameIndex = mainWb.Worksheets.Names.Add("MyExternalRange");
            Name myExternalRange = mainWb.Worksheets.Names[nameIndex];

            // Use SetRefersTo to define the reference (A1 style, not locale‑specific)
            myExternalRange.SetRefersTo("='[ExternalData.xlsx]Sheet1'!$A$1:$A$3", false, false);

            // ---------------------------------------------------------------
            // Step 5: Use the named range in a formula inside the main workbook
            // ---------------------------------------------------------------
            mainSheet.Cells["B1"].Formula = "=SUM(MyExternalRange)";

            // ---------------------------------------------------------------
            // Step 6: Calculate formulas, providing the external workbook as a linked data source
            // ---------------------------------------------------------------
            CalculationOptions calcOptions = new CalculationOptions();
            calcOptions.LinkedDataSources = new Workbook[] { externalWb };
            mainWb.CalculateFormula(calcOptions);

            // Output the result to the console
            Console.WriteLine("Result of SUM(MyExternalRange): " + mainSheet.Cells["B1"].Value);

            // ---------------------------------------------------------------
            // Step 7: Save the main workbook (lifecycle rule: use provided save)
            // ---------------------------------------------------------------
            mainWb.Save("MainWorkbookWithExternalNamedRange.xlsx");
        }
    }
}
