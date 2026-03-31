using System;
using System.Globalization;
using Aspose.Cells;

class LoadWorkbookWithEnglishFormulas
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Create LoadOptions and set the culture to English (US)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("en-US");
        // Ensure formulas are parsed when the workbook is opened
        loadOptions.ParsingFormulaOnOpen = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Calculate all formulas to verify correct evaluation under English locale
        workbook.CalculateFormula();

        // Save the workbook (optional) to confirm formulas remain in English syntax
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}