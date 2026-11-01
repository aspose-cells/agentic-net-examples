// Title: Loop to add worksheets, assign locale-based names, and set workbook-wide GlobalizationSettings for subtotal labels using Aspose.Cells for .NET
// Common Searches: how to add multiple worksheets in a loop and set workbook culture with Aspose.Cells C# | assign language-specific worksheet names while configuring GlobalizationSettings in Aspose.Cells | set subtotal label language for an Excel workbook using Aspose.Cells .NET API | create workbook with different locale codes in sheet names using Aspose.Cells
// Tags: loop add worksheets Aspose.Cells C# | workbook-wide globalization settings Aspose.Cells | locale-based worksheet naming Aspose.Cells | configure subtotal label language .NET Excel | custom CultureInfo for Aspose.Cells workbook

using System;
using System.Globalization;
using Aspose.Cells;

// The example creates a new Workbook, sets the workbook’s CultureInfo to en-US, adds four additional worksheets in a loop (total five), assigns each sheet a name that includes a locale code (en-US, fr-FR, de-DE, es-ES, it-IT), and saves the file as CustomGlobalizationSettings.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the global culture for the workbook (affects subtotal labels, dates, numbers, etc.)
            workbook.Settings.CultureInfo = new CultureInfo("en-US");

            // Ensure the workbook has enough worksheets (total 5)
            for (int i = 1; i < 5; i++)
            {
                workbook.Worksheets.Add();
            }

            // Optional: give each worksheet a distinct name for clarity
            string[] languages = { "en-US", "fr-FR", "de-DE", "es-ES", "it-IT" };
            for (int i = 0; i < workbook.Worksheets.Count && i < languages.Length; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Name = $"Sheet_{i + 1}_{languages[i]}";
            }

            // Save the workbook to a file
            string outputPath = "CustomGlobalizationSettings.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
