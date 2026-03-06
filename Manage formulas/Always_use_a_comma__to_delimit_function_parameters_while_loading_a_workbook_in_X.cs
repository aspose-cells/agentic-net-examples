using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // If the source file does not exist, create a simple workbook for demonstration
        if (!File.Exists(sourcePath))
        {
            // Create a new workbook and add a simple formula that uses commas
            Workbook tempWb = new Workbook();
            tempWb.Worksheets[0].Cells["A1"].Formula = "=SUM(1,2,3)";
            tempWb.Save(sourcePath);
        }

        // Create LoadOptions (default constructor) – follows the provided rule
        LoadOptions loadOptions = new LoadOptions();

        // Load the workbook with the specified LoadOptions – follows the provided rule
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Ensure that the list separator for function parameters is a comma
        SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings();
        globalizationSettings.SetListSeparator(','); // explicitly set comma as separator
        workbook.Settings.GlobalizationSettings = globalizationSettings;

        // Demonstrate the effect by adding a formula that uses the comma separator
        workbook.Worksheets[0].Cells["B1"].Formula = "=CONCATENATE(\"Hello\",\" World\")";

        // Calculate formulas to ensure they are evaluated correctly
        workbook.CalculateFormula();

        // Save the modified workbook – follows the provided rule
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine("Workbook loaded and saved with comma as list separator.");
    }
}