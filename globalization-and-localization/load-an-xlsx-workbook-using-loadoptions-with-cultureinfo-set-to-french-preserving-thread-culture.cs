// Title: How to load an XLSX workbook with French (fr-FR) CultureInfo using Aspose.Cells LoadOptions while preserving the original thread culture in C#
// Common Searches: C# load Excel file with French locale using Aspose.Cells LoadOptions | preserve original thread culture when changing CultureInfo for Aspose.Cells workbook loading | example of Aspose.Cells LoadOptions setting CultureInfo to fr-FR | set thread culture to fr-FR before loading an XLSX workbook in C# | restore thread culture after loading an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells LoadOptions CultureInfo fr-FR | load XLSX workbook with specific locale | preserve thread culture during Excel load | C# globalization Aspose.Cells | Excel file loading culture settings

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The code saves the current thread culture, switches the thread to French (fr-FR), configures LoadOptions with the same CultureInfo, loads an XLSX workbook using those options, optionally processes the workbook, saves it, and finally restores the original thread culture.
class Program
{
    static void Main()
    {
        // Preserve the original thread culture
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

        try
        {
            // Set the thread culture to French (France)
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

            // Configure LoadOptions with French culture
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = new CultureInfo("fr-FR")
            };

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the XLSX workbook using the configured LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Process the workbook here

            // Save the workbook if needed
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as appropriate
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Restore the original thread culture
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }
    }
}
