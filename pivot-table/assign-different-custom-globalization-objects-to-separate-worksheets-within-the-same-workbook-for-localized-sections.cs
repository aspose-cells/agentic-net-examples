// Title: C# – Apply Different GlobalizationSettings to Individual Worksheets in Aspose.Cells
// Description: Demonstrates how to assign custom GlobalizationSettings objects to separate worksheets in a single Aspose.Cells workbook, localizing boolean and error strings for English and Russian sheets, saving the file, and reading back the localized values.
// Keywords: Aspose.Cells | C# | .NET | GlobalizationSettings | per worksheet localization | custom globalization | English worksheet | Russian worksheet | localized boolean strings | localized error values | Excel workbook localization | save workbook Aspose.Cells | read localized cells
// Common Searches: Aspose.Cells set GlobalizationSettings per sheet | C# localize Excel worksheet with Aspose.Cells | how to display Russian boolean values in Aspose.Cells | custom error messages for different worksheets Aspose.Cells | save mixed‑language workbook using Aspose.Cells
// Developer Intent: Create a workbook where each worksheet uses its own GlobalizationSettings to show language‑specific boolean and error strings.
// Use Cases: Generate an English sheet with default strings and a Russian sheet with RussianGlobalizationSettings, then save both in one workbook. | Load the saved workbook and output the localized strings for each sheet to confirm per‑sheet globalization. | Add additional language worksheets (e.g., French, Spanish) by defining new GlobalizationSettings subclasses and applying them without affecting existing sheets.
// AI Prompts: Write C# code that assigns a custom GlobalizationSettings object to a specific worksheet in Aspose.Cells while leaving other sheets unchanged. | Show how to create a FrenchGlobalizationSettings class and apply it to a new worksheet in an existing Aspose.Cells workbook. | Explain how to retrieve and display localized boolean and error strings from a workbook that used different GlobalizationSettings per worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Custom globalization for English (default behavior)
    // Demonstrates how to assign custom GlobalizationSettings objects to separate worksheets in a single Aspose.Cells workbook, localizing boolean and error strings for English and Russian sheets, saving the file, and reading back the localized values.
    public class EnglishGlobalizationSettings : GlobalizationSettings
    {
        // No overrides – keep the default English strings
    }

    // Custom globalization for Russian
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Override boolean display strings
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Override error value strings
        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#DIV/0!" => "#ДЕЛ/0!",
                "#NAME?"  => "#ИМЯ?",
                "#REF!"   => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A"    => "#Н/Д",
                "#NUM!"   => "#ЧИСЛО!",
                "#NULL!"  => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }

        // Override default sheet name (illustrative)
        public override string GetDefaultSheetName()
        {
            return "Лист";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add two worksheets for demonstration
                Worksheet sheetEn = workbook.Worksheets[0]; // First sheet (default name)
                Worksheet sheetRu = workbook.Worksheets.Add("Russian");

                // Populate English sheet using English globalization settings
                workbook.Settings.GlobalizationSettings = new EnglishGlobalizationSettings();

                Cells cellsEn = sheetEn.Cells;
                cellsEn[0, 0].PutValue(true);   // Boolean true
                cellsEn[0, 1].PutValue(false);  // Boolean false
                cellsEn[0, 2].PutValue("#DIV/0!"); // Error value

                // Populate Russian sheet using Russian globalization settings
                workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

                Cells cellsRu = sheetRu.Cells;
                cellsRu[0, 0].PutValue(true);   // Will display "ИСТИНА"
                cellsRu[0, 1].PutValue(false);  // Will display "ЛОЖЬ"
                cellsRu[0, 2].PutValue("#DIV/0!"); // Will display "#ДЕЛ/0!"

                // Save the workbook (lifecycle rule: create → save)
                string outputPath = "LocalizedWorksheets.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");

                // Demonstrate reading back the localized strings (optional)
                if (File.Exists(outputPath))
                {
                    Workbook loadedWb = new Workbook(outputPath);

                    // English sheet (first sheet) – uses default globalization (English)
                    Console.WriteLine("\nEnglish sheet values:");
                    Console.WriteLine(loadedWb.Worksheets[0].Cells[0, 0].StringValue); // TRUE
                    Console.WriteLine(loadedWb.Worksheets[0].Cells[0, 1].StringValue); // FALSE
                    Console.WriteLine(loadedWb.Worksheets[0].Cells[0, 2].StringValue); // #DIV/0!

                    // Russian sheet (second sheet) – uses Russian globalization that was active
                    // when the values were written, so the stored strings are already localized.
                    Console.WriteLine("\nRussian sheet values:");
                    Console.WriteLine(loadedWb.Worksheets[1].Cells[0, 0].StringValue); // ИСТИНА
                    Console.WriteLine(loadedWb.Worksheets[1].Cells[0, 1].StringValue); // ЛОЖЬ
                    Console.WriteLine(loadedWb.Worksheets[1].Cells[0, 2].StringValue); // #ДЕЛ/0!
                }
                else
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
