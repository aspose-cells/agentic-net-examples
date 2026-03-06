using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook. ParsingFormulaOnOpen is set to false to speed up loading.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Set a low limit for the maximum rows of a shared formula.
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Work with the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Attempt to set a shared formula that exceeds the current limit (101 rows).
        // This demonstrates how the limit affects shared formula creation.
        cells["B1"].SetSharedFormula("=A1", 101, 1);

        // Display the formula in the last cell of the range to verify behavior.
        Console.WriteLine("Formula in B101: " + cells["B101"].Formula);

        // Increase the limit to a higher value and repeat the operation on a new sheet.
        workbook.Settings.MaxRowsOfSharedFormula = 1024;

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        Cells cells2 = sheet2.Cells;
        cells2["B1"].SetSharedFormula("=A1", 101, 1);
        Console.WriteLine("Formula in B101 (Sheet2): " + cells2["B101"].Formula);

        // Save the workbook with the modifications.
        workbook.Save("output.xlsx");
    }
}