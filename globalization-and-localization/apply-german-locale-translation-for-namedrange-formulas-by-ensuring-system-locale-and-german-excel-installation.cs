// Title: Apply German (de-DE) locale to named‑range formulas and recalculate workbook with Aspose.Cells for .NET (C#)
// Common Searches: how to set workbook culture to de-DE in Aspose.Cells C# | recalculate Excel formulas using German locale with Aspose.Cells .NET | Aspose.Cells example for translating named range formulas to German | set thread culture to German for Excel processing in C# using Aspose.Cells | convert Excel workbook formulas to German language with Aspose.Cells
// Tags: workbook culture de-DE Aspose.Cells | formula recalculation German locale Aspose.Cells | named range translation Aspose.Cells | thread culture de-DE C# Excel processing | save workbook with German settings Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program loads an existing Excel file, sets the current thread and the workbook's Settings.CultureInfo to the German locale (de-DE), recalculates all formulas—including those in named ranges—under this culture, and saves the result as a new workbook. It includes checks for missing input files and generic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_de.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Simulate German locale for the current thread
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("de-DE");

                // Apply German locale to the workbook for formula translation
                workbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Recalculate formulas using the German locale
                workbook.CalculateFormula();

                // Save the workbook with the desired output name
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
