// Title: Load an Excel workbook with a custom French (fr-FR) culture using Aspose.Cells for .NET
// Common Searches: Aspose.Cells load workbook with specific CultureInfo in C# | set workbook.Settings.CultureInfo before opening Excel file using Aspose.Cells | apply French (fr-FR) locale to Aspose.Cells workbook loading | configure thread CurrentCulture for Aspose.Cells Excel import in .NET | load Excel file with custom globalization settings using Aspose.Cells .NET
// Tags: load workbook custom CultureInfo Aspose.Cells | Aspose.Cells workbook.Settings.CultureInfo usage | French locale Aspose.Cells Excel loading | globalization settings thread culture Aspose.Cells | C# Aspose.Cells workbook loading culture

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;
using System.Threading;

// The example shows how to define a French (fr-FR) CultureInfo, assign it to the current thread, verify the Excel file exists, load the file into an Aspose.Cells Workbook, set the same CultureInfo on the workbook's Settings, and properly dispose the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Define a custom culture (e.g., French - France)
            CultureInfo customCulture = new CultureInfo("fr-FR");

            // Apply the custom culture to the current thread (globalization settings)
            Thread.CurrentThread.CurrentCulture = customCulture;
            Thread.CurrentThread.CurrentUICulture = customCulture;

            // Path to the input Excel file
            string inputPath = "InputWorkbook.xlsx";

            // Ensure the file exists before attempting to load it
            if (File.Exists(inputPath))
            {
                Workbook workbook = null;

                try
                {
                    // Load the Excel file into the workbook using the appropriate constructor
                    workbook = new Workbook(inputPath);
                    // Assign the custom culture to the workbook settings
                    workbook.Settings.CultureInfo = customCulture;

                    // Further processing can be done here...
                    Console.WriteLine("Workbook loaded successfully.");
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                }
                finally
                {
                    // Dispose workbook if needed (Aspose.Cells implements IDisposable in newer versions)
                    if (workbook is IDisposable disposableWorkbook)
                    {
                        disposableWorkbook.Dispose();
                    }
                }
            }
            else
            {
                Console.WriteLine($"Input file not found: {inputPath}");
            }
        }
        catch (Exception ex)
        {
            // Handle unexpected exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
