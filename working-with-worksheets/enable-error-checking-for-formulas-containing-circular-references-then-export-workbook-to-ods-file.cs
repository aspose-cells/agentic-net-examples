// Title: Handle Circular References with Iterative Calculation and Export to ODS using Aspose.Cells for .NET (C#)
// Description: Create a workbook, define a circular reference between A1 and B1, enable iterative calculation (max iterations & change), calculate formulas, set OdsSaveOptions (LibreOffice generator, ODF 1.2), and save the file as ODS.
// Keywords: Aspose.Cells | circular reference | iterative calculation | C# | .NET | ODS export | OdsSaveOptions | LibreOffice generator | ODF 1.2 | formula calculation
// Common Searches: Aspose.Cells circular reference handling | Enable iterative calculation Aspose.Cells .NET | Save workbook as ODS with Aspose.Cells | Configure OdsSaveOptions C# | Calculate formulas with circular dependencies Aspose
// Developer Intent: Resolve circular references through iterative calculation and then export the workbook to an ODS document.
// Use Cases: Financial models with inter‑dependent cells that must be calculated before generating an ODS report. | Engineering spreadsheets containing feedback loops, exported to ODS for cross‑platform sharing. | Automated batch processing that validates circular formulas and produces ODF‑1.2‑compliant files.
// AI Prompts: Generate C# code using Aspose.Cells that enables iterative calculation for circular references and saves the workbook as an ODS file with the LibreOffice generator. | Show how to set MaxIteration and MaxChange in FormulaSettings, run CalculateFormula, and configure OdsSaveOptions for ODF 1.2 compliance. | Explain step‑by‑step how to detect circular references, apply iterative calculation, and export the result to ODS with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Create a workbook, define a circular reference between A1 and B1, enable iterative calculation (max iterations & change), calculate formulas, set OdsSaveOptions (LibreOffice generator, ODF 1.2), and save the file as ODS.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define a circular reference: A1 depends on B1 and B1 depends on A1
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve the circular reference
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

        // Perform formula calculation
        workbook.CalculateFormula();

        // Set ODS save options (optional customizations)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Export the workbook to an ODS file
        workbook.Save("CircularReferenceDemo.ods", saveOptions);
    }
}
