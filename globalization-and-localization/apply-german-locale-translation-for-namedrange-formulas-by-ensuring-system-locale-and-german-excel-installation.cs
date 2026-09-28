// Title: Apply German (de-DE) locale to Excel workbook formulas with Aspose.Cells for .NET
// AI Prompts: Set the current thread's CultureInfo to de‑DE, load the workbook, assign workbook.Settings.CultureInfo to the German culture, and save the file. | Configure Aspose.Cells to parse formulas using German regional settings by updating the workbook's culture before exporting.
// Common Searches: Aspose.Cells set German de-DE culture for formula evaluation in C# | How to localize Excel formulas to German using Aspose.Cells .NET | C# load workbook and apply German locale for named‑range formulas with Aspose | Change formula parsing language to German in Aspose.Cells workbook | Set workbook.Settings.CultureInfo to de-DE for Excel localization in .NET
// Tags: set workbook cultureinfo de-de Aspose.Cells | German formula locale Aspose.Cells .NET | C# Excel workbook localization Aspose | apply regional settings to workbook formulas | Aspose.Cells cultureinfo configuration

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example sets the thread's culture to German (de-DE), loads an input.xlsx workbook, assigns the German CultureInfo to workbook.Settings.CultureInfo for locale‑specific formula handling, and saves the modified workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Set thread culture to German (de-DE)
                CultureInfo germanCulture = new CultureInfo("de-DE");
                Thread.CurrentThread.CurrentCulture = germanCulture;
                Thread.CurrentThread.CurrentUICulture = germanCulture;

                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Apply German locale for formula parsing
                workbook.Settings.CultureInfo = germanCulture;
                // Note: UseFormulaLocale property is not available in the current Aspose.Cells version.
                // Setting CultureInfo is sufficient for locale‑specific formula handling.

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
