using System;
using Aspose.Cells;

class SharedFormulaDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with sample values
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1); // A1‑A5 = 1‑5
        }

        // Set a shared formula in column B (B1:B5) that multiplies column A by 2
        cells["B1"].SetSharedFormula("=A1*2", 5, 1);

        // Save the workbook to a file
        string filePath = "SharedFormulaDemo.xlsx";
        workbook.Save(filePath);

        // Load the workbook with LoadOptions that skip formula parsing on open
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = false; // keep formulas unparsed initially
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

        // Parse all formulas after loading
        loadedWorkbook.ParseFormulas(false);

        // Output the values of columns A and B to verify the shared formula worked
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
        for (int i = 0; i < 5; i++)
        {
            Console.WriteLine($"A{i + 1} = {loadedSheet.Cells[i, 0].Value}, B{i + 1} = {loadedSheet.Cells[i, 1].Value}");
        }
    }
}