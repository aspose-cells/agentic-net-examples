using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Create load options and ensure formulas are parsed when the workbook is opened
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Create a settable globalization settings instance
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Force the list separator to be a comma (',')
        globalization.SetListSeparator(',');

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Re‑parse all formulas so they use the new list separator
        workbook.ParseFormulas(false);

        // Save the workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}