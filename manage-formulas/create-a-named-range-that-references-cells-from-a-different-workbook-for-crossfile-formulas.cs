// Title: Aspose.Cells .NET: Define a Named Range that References an External Workbook
// Description: This example shows how to create a secondary workbook, save it, add an external link to a primary workbook, define an external name, and then create a named range in the primary workbook that points to the external range using the syntax ='[ExternalData.xlsx]DataSheet'!$A$1:$A$3. The named range is used in a SUM formula, calculation options are set with LinkedDataSources, the formula is evaluated, and both workbooks are saved.
// Keywords: Aspose.Cells | .NET | C# | external named range | cross workbook formula | Excel external reference | linked data sources | calculate external formulas | named range syntax | external link Aspose.Cells
// Common Searches: Aspose.Cells create named range to external workbook | C# external link formula Aspose.Cells | how to sum range from another Excel file using Aspose.Cells | set LinkedDataSources for external workbook calculation | reference cells in another file with Aspose.Cells
// Developer Intent: Create a named range in a workbook that points to a range in a different workbook and use it in a formula.
// Use Cases: Build a reporting workbook that aggregates values from a shared data file without opening the source file. | Perform financial calculations on quarterly figures stored in separate workbooks via a single named range. | Consolidate sensor data from multiple Excel files into a master sheet for real‑time analysis.
// AI Prompts: Generate C# code using Aspose.Cells to create an external workbook, add an external link, define a named range that references the external range, and calculate a SUM formula. | Explain how to configure CalculationOptions.LinkedDataSources to include external workbooks for formula evaluation in Aspose.Cells. | Provide steps to update the reference of a cross‑file named range when the source workbook name or sheet name changes.

using System;
using Aspose.Cells;

// This example shows how to create a secondary workbook, save it, add an external link to a primary workbook, define an external name, and then create a named range in the primary workbook that points to the external range using the syntax ='[ExternalData.xlsx]DataSheet'!$A$1:$A$3. The named range is used in a SUM formula, calculation options are set with LinkedDataSources, the formula is evaluated, and both workbooks are saved.
class CrossFileNamedRangeDemo
{
    static void Main()
    {
        // ---------- Create external workbook ----------
        Workbook externalWb = new Workbook();
        Worksheet extSheet = externalWb.Worksheets[0];
        extSheet.Name = "DataSheet";

        // Populate some data in the external workbook
        extSheet.Cells["A1"].PutValue(10);
        extSheet.Cells["A2"].PutValue(20);
        extSheet.Cells["A3"].PutValue(30);

        // Save the external workbook to a physical file (required for external linking)
        string externalPath = "ExternalData.xlsx";
        externalWb.Save(externalPath);

        // ---------- Create main workbook ----------
        Workbook mainWb = new Workbook();
        Worksheet mainSheet = mainWb.Worksheets[0];
        mainSheet.Name = "MainSheet";

        // Add an external link that points to the external workbook
        string[] sheetNames = new string[] { extSheet.Name };
        int linkIndex = mainWb.Worksheets.ExternalLinks.Add(externalPath, sheetNames);
        ExternalLink extLink = mainWb.Worksheets.ExternalLinks[linkIndex];

        // Add an external name inside the external link (optional, shows how to name the range)
        extLink.AddExternalName("ExtRange", "=DataSheet!$A$1:$A$3");

        // Create a named range in the main workbook that references the external range
        int nameIdx = mainWb.Worksheets.Names.Add("CrossFileRange");
        Name crossName = mainWb.Worksheets.Names[nameIdx];
        // The RefersTo string uses the external reference syntax
        crossName.RefersTo = "='[ExternalData.xlsx]DataSheet'!$A$1:$A$3";

        // Use the named range in a formula inside the main workbook
        mainSheet.Cells["B1"].Formula = "=SUM(CrossFileRange)";

        // Configure calculation options to include the external workbook as a linked data source
        CalculationOptions calcOptions = new CalculationOptions();
        calcOptions.LinkedDataSources = new Workbook[] { externalWb };

        // Calculate formulas using the configured options
        mainWb.CalculateFormula(calcOptions);

        // Output the result of the cross‑file formula
        Console.WriteLine("Sum of external range: " + mainSheet.Cells["B1"].Value);

        // Save the main workbook
        mainWb.Save("MainWithCrossFileNamedRange.xlsx");
    }
}
