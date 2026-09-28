// Title: Preserve the original thread culture while loading an Excel workbook with French CultureInfo using Aspose.Cells LoadOptions in C#
// AI Prompts: Load an XLSX workbook with a French CultureInfo via Aspose.Cells LoadOptions, temporarily set Thread.CurrentThread.CurrentCulture to French, and then restore the original culture after the workbook is loaded. | Add robust error handling that guarantees the original thread culture is reinstated even when the input file is missing or an exception occurs during workbook loading.
// Common Searches: c# aspocells load workbook with french cultureinfo without changing global thread culture | how to use LoadOptions.CultureInfo to load an Excel file in a specific locale | restore original thread culture after Aspose.Cells workbook load | aspocells example for temporary thread culture switch during file load | check file existence before loading workbook with Aspose.Cells in C#
// Tags: load workbook with specific CultureInfo Aspose.Cells | temporary thread culture switch C# | Aspose.Cells LoadOptions CultureInfo setting | handle missing Excel file Aspose.Cells | restore original thread culture after workbook load

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example stores the current thread culture, sets a French CultureInfo on LoadOptions, temporarily changes Thread.CurrentThread.CurrentCulture to French before loading the workbook, restores the original culture after loading (including in catch blocks), validates the input file's existence, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Store the original thread culture to restore it later.
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

            // Define the French culture that will be used while loading the workbook.
            CultureInfo frenchCulture = new CultureInfo("fr-FR");

            // Create LoadOptions and assign the French culture to it.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = frenchCulture
            };

            // Path to the input workbook.
            string inputPath = "SampleFrench.xlsx";

            try
            {
                // Prevent FileNotFoundException by checking file existence.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Temporarily switch the thread culture to French before loading.
                Thread.CurrentThread.CurrentCulture = frenchCulture;

                // Load the workbook using the specified LoadOptions.
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Restore the original thread culture after loading.
                Thread.CurrentThread.CurrentCulture = originalCulture;

                // Save the workbook to verify successful loading.
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Ensure original culture is restored in case of an error.
                Thread.CurrentThread.CurrentCulture = originalCulture;
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
