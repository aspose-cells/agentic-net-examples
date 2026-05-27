using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create a circular reference scenario
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve the circular reference
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // maximum change threshold

        // Perform formula calculation
        workbook.CalculateFormula();

        // Set ODS save options (optional customizations)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Export the workbook to an ODS file
        workbook.Save("CircularReferenceDemo.ods", saveOptions);
    }
}