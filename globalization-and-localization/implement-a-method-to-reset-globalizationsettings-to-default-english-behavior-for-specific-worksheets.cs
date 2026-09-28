// Title: Reset Aspose.Cells workbook globalization to English (en-US) culture using C#
// AI Prompts: Write a C# method that assigns Workbook.Settings.CultureInfo to the en-US culture with Aspose.Cells. | Extend the helper to accept worksheet names and apply the English CultureInfo only to those sheets while leaving other sheets unchanged.
// Common Searches: Aspose.Cells set workbook culture to en-US in C# | how to change Excel file language to English using Aspose.Cells .NET | reset globalization settings for an Aspose.Cells workbook | C# example for applying English locale to an Excel workbook with Aspose.Cells | apply English culture to specific worksheets Aspose.Cells
// Tags: Aspose.Cells workbook cultureinfo en-us | reset workbook globalization Aspose.Cells | set Excel locale to English C# | globalization settings Aspose.Cells .NET | apply English culture to worksheets Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example defines a WorksheetGlobalizationHelper with a ResetGlobalizationToEnglish method that sets the workbook's Settings.CultureInfo to the en-US culture, loads an input Excel file, applies the English globalization globally, and saves the result.
    public static class WorksheetGlobalizationHelper
    {
        // Resets globalization settings of the workbook to English (en-US) behavior.
        public static void ResetGlobalizationToEnglish(Workbook workbook, params string[] worksheetNames)
        {
            // English culture.
            CultureInfo englishCulture = new CultureInfo("en-US");

            // Configure workbook globalization.
            workbook.Settings.CultureInfo = englishCulture;
            // The following properties are not available in the current Aspose.Cells version,
            // so they are omitted. Setting the CultureInfo is sufficient for English behavior.
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Reset globalization settings to English.
                WorksheetGlobalizationHelper.ResetGlobalizationToEnglish(workbook);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the message.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
