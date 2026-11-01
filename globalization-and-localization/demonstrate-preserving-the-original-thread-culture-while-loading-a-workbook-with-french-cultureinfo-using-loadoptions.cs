// Title: Preserve original thread culture while loading an Excel workbook with French CultureInfo using Aspose.Cells LoadOptions in C#
// Common Searches: c# load excel workbook with specific cultureinfo using aspose.cells loadoptions | preserve thread culture after switching to French locale in Aspose.Cells | set French culture for workbook loading without affecting global thread settings | how to restore original CultureInfo after loading an Excel file with Aspose.Cells | use LoadOptions.CultureInfo to read localized numbers in Excel with Aspose.Cells
// Tags: loadoptions cultureinfo excel c# | aspose.cells thread culture management | french locale workbook loading | cultureinfo reset after workbook load | excel globalization with aspose.cells

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;

// The example saves the current thread's CultureInfo, switches the thread to French (fr-FR), verifies the Excel file exists, configures LoadOptions with French CultureInfo, loads the workbook, reads a cell value, and finally restores the original thread cultures while handling any exceptions.
class Program
{
    static void Main()
    {
        // Preserve original thread cultures
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        CultureInfo originalUICulture = Thread.CurrentThread.CurrentUICulture;

        try
        {
            // Switch thread culture to French (France)
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-FR");

            string filePath = "SampleFr.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return;
            }

            // Configure LoadOptions with French culture
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = new CultureInfo("fr-FR")
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example operation: read a cell value
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Value in A1: {cellValue}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Restore the original thread cultures
            Thread.CurrentThread.CurrentCulture = originalCulture;
            Thread.CurrentThread.CurrentUICulture = originalUICulture;
        }
    }
}
