// Title: Evaluate External Workbook References with Worksheet.CalculateFormula in Aspose.Cells for .NET
// Description: Shows how to create an in‑memory external workbook, place a value in Sheet1!A2, set a formula like =[External.xlsx]Sheet1!$A$2 in a main workbook, configure CalculationOptions.LinkedDataSources, and call Worksheet.CalculateFormula to obtain the result without persisting the source file.
// Keywords: Aspose.Cells | Worksheet.CalculateFormula | external workbook reference | CalculationOptions.LinkedDataSources | C# | .NET | cross‑workbook formula | in‑memory workbook | evaluate formula
// Common Searches: Aspose.Cells calculate formula from another workbook | Worksheet.CalculateFormula external reference example | How to use CalculationOptions.LinkedDataSources in C# | Evaluate cross‑workbook formula without saving file Aspose.Cells | C# Aspose.Cells external cell reference calculation
// Developer Intent: Execute a formula that reads a cell from a different workbook by providing the source workbook through CalculationOptions.LinkedDataSources.
// Use Cases: Perform on‑the‑fly calculations for financial models where data resides in separate workbook objects. | Generate a report that aggregates values from a template workbook without writing intermediate files. | Validate external data links during automated testing of spreadsheet‑based workflows.
// AI Prompts: Provide a C# example that uses Worksheet.CalculateFormula with an external workbook reference and CalculationOptions.LinkedDataSources. | Show how to link multiple external workbooks in Aspose.Cells and evaluate formulas that reference them. | Explain error handling for missing or corrupted external workbooks when using Worksheet.CalculateFormula.

using System;
using Aspose.Cells;

namespace AsposeCellsExternalFormulaDemo
{
    // Shows how to create an in‑memory external workbook, place a value in Sheet1!A2, set a formula like =[External.xlsx]Sheet1!$A$2 in a main workbook, configure CalculationOptions.LinkedDataSources, and call Worksheet.CalculateFormula to obtain the result without persisting the source file.
    class Program
    {
        static void Main()
        {
            // ---------- Create external workbook ----------
            // This workbook will act as the data source for the external reference.
            Workbook externalWb = new Workbook();
            // Put a sample value in Sheet1!A2 (the cell we will reference).
            externalWb.Worksheets[0].Cells["A2"].PutValue(12345);

            // ---------- Create main workbook ----------
            Workbook mainWb = new Workbook();
            Worksheet sheet = mainWb.Worksheets[0];

            // Formula that references a cell in the external workbook.
            // The external workbook is identified by its file name "External.xlsx".
            // In this demo we don't actually save the external workbook to disk;
            // we link it via CalculationOptions.LinkedDataSources.
            string externalFormula = "=[External.xlsx]Sheet1!$A$2";

            // Set the formula in cell A1 of the main worksheet.
            sheet.Cells["A1"].Formula = externalFormula;

            // ---------- Prepare calculation options ----------
            // Link the external workbook so that the calculation engine can resolve the reference.
            CalculationOptions calcOptions = new CalculationOptions
            {
                // The array can contain multiple external workbooks if needed.
                LinkedDataSources = new Workbook[] { externalWb }
            };

            // ---------- Calculate the formula ----------
            // Use Worksheet.CalculateFormula overload that accepts a formula string and options.
            // This returns the evaluated result directly.
            object result = sheet.CalculateFormula(externalFormula, calcOptions);

            // Output the result.
            Console.WriteLine($"Result of formula '{externalFormula}' = {result}");

            // (Optional) Save the main workbook to verify the calculated value is stored.
            mainWb.Save("MainWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
