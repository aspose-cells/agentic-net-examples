// Title: Toggle Excel formulas between English and localized names using Formula and FormulaLocal in Aspose.Cells for .NET
// AI Prompts: Write C# code that sets Workbook.Settings.CultureInfo to a specific locale, assigns a formula with FormulaLocal, and then reads the equivalent English formula via the Formula property. | Demonstrate how assigning an English formula string to the FormulaLocal property automatically converts the function name to the workbook’s current language. | Create a sample that switches a formula from English to Portuguese (or another locale) and back, then saves the workbook as an XLSX file.
// Common Searches: Aspose.Cells how to use FormulaLocal to write Portuguese function names | retrieve English version of a localized Excel formula with Aspose.Cells C# | set workbook culture info for localized formulas in Aspose.Cells .NET | convert SUM formula to SOMA automatically using FormulaLocal Aspose.Cells | toggle between English and localized Excel formulas programmatically
// Tags: Aspose.Cells FormulaLocal localization | C# set workbook CultureInfo Aspose.Cells | convert Excel function name to locale Aspose.Cells | retrieve English formula from localized cell | automatic formula language conversion Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, sets its culture to Portuguese (Brazil), writes formulas using both English (SUM) and localized (SOMA) names, shows how to read each version via the Formula and FormulaLocal properties, demonstrates that assigning an English formula to FormulaLocal automatically translates it to the locale’s function name, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the workbook culture to Portuguese (Brazil) to demonstrate localized function names
            workbook.Settings.CultureInfo = new CultureInfo("pt-BR");

            Worksheet sheet = workbook.Worksheets[0];

            // 1. Use English function name (SUM) – stored in the Formula property
            sheet.Cells["A1"].Formula = "=SUM(10,20)";

            // 2. Use localized function name (SOMA) – stored in the FormulaLocal property
            // Note: In Portuguese the argument separator is ';'
            sheet.Cells["A2"].FormulaLocal = "=SOMA(10;20)";

            // 3. Retrieve the English formula from a cell that originally used a localized formula
            string englishFromLocal = sheet.Cells["A2"].Formula; // Returns "=SUM(10,20)"

            // 4. Retrieve the localized formula from a cell that originally used an English formula
            string localFromEnglish = sheet.Cells["A1"].FormulaLocal; // Returns "=SOMA(10;20)" in the current locale

            // 5. Toggle: assign the English formula to a new cell using FormulaLocal
            // Aspose.Cells automatically converts the function name to the localized version
            sheet.Cells["A3"].FormulaLocal = englishFromLocal; // Cell A3 will contain "=SOMA(10;20)"

            // 6. Toggle: assign the localized formula to a new cell using Formula (English)
            // Use the previously obtained English formula to avoid separator issues
            sheet.Cells["A4"].Formula = englishFromLocal; // Cell A4 will contain "=SUM(10,20)"

            // Define output file path
            string outputPath = "ToggleFormulaLocal.xlsx";

            // Ensure the directory exists (in case a relative path is used)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
