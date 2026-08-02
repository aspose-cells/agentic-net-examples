// Title: C# – Handle Circular References with Iterative Calculation and Export to ODS using Aspose.Cells
// Description: Demonstrates how to enable iterative calculation for circular references, run formula evaluation, configure OdsSaveOptions (LibreOffice generator), and save the workbook as an ODS file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# circular reference | iterative calculation Aspose.Cells | save workbook as ODS | OdsSaveOptions LibreOffice | formula calculation settings
// Common Searches: Aspose.Cells enable iterative calculation | circular reference handling in .NET Excel | export Aspose.Cells workbook to ODS | set OdsSaveOptions generator type | calculate formulas with circular references
// Developer Intent: Enable iterative calculation to resolve circular references and save the workbook in ODS format.
// Use Cases: Process worksheets that contain mutually dependent cells by activating iterative calculation with custom iteration limits and convergence thresholds. | Export a workbook with circular formulas to an ODS file compatible with LibreOffice or OpenOffice. | Fine‑tune MaxIteration and MaxChange to control precision before saving the document.
// AI Prompts: Show how to detect circular references in Aspose.Cells without using iterative calculation and throw an exception. | Provide code that saves a workbook to ODS while preserving cell styles, comments, and hyperlinks. | Explain how to log each iteration step when EnableIterativeCalculation is true in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Demonstrates how to enable iterative calculation for circular references, run formula evaluation, configure OdsSaveOptions (LibreOffice generator), and save the workbook as an ODS file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create a circular reference: A1 depends on B1 and B1 depends on A1
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

        // Perform formula calculation
        workbook.CalculateFormula();

        // Set ODS save options (optional customizations)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Export the workbook to an ODS file
        workbook.Save("CircularReferenceDemo.ods", saveOptions);
    }
}
