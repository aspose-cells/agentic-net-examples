// Title: Create multiple worksheets with individual GlobalizationSettings for localized subtotal labels using a C# loop in Aspose.Cells
// AI Prompts: Generate C# code that iterates over a list of culture identifiers, adds or reuses worksheets, and assigns a new GlobalizationSettings object to each worksheet so that subtotal labels appear in the worksheet's language. | Show how to set Worksheet.CustomGlobalizationSettings.CultureInfo inside a loop while also naming each sheet with its locale code. | Provide a complete Aspose.Cells example that saves the workbook after configuring per‑sheet globalization and prints the output file path. | Explain how to customize the SubtotalLabel property of GlobalizationSettings for each worksheet based on its culture.
// Common Searches: asp.net aspose.cells assign different GlobalizationSettings per worksheet | how to localize subtotal labels for each sheet in Excel using Aspose.Cells C# | loop create worksheets with culture-specific names and settings Aspose.Cells | per sheet culture info for subtotal functions Aspose.Cells .NET example | custom globalization settings for multiple worksheets in Aspose.Cells
// Tags: per‑worksheet GlobalizationSettings Aspose.Cells | custom subtotal label language C# | loop assign culture to Excel worksheets Aspose.Cells | localized worksheet naming .NET | Excel subtotal localization Aspose.Cells | worksheet.CustomGlobalizationSettings usage

using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // The example creates a Workbook, loops through a list of culture codes (en-US, fr-FR, de-DE, es-ES), adds or reuses worksheets, names each sheet with the corresponding locale, creates a distinct GlobalizationSettings object for each worksheet, assigns it to Worksheet.CustomGlobalizationSettings to localize subtotal labels, and finally saves the workbook as 'GlobalizationSettingsDemo.xlsx'.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Languages used for worksheet naming and optional global culture setting
                List<string> languages = new List<string> { "en-US", "fr-FR", "de-DE", "es-ES" };

                // Create or reuse worksheets and assign names
                for (int i = 0; i < languages.Count; i++)
                {
                    Worksheet sheet;
                    if (i < workbook.Worksheets.Count)
                    {
                        // Reuse existing worksheet at the given index
                        sheet = workbook.Worksheets[i];
                    }
                    else
                    {
                        // Add a new worksheet and retrieve its reference
                        int newIndex = workbook.Worksheets.Add();
                        sheet = workbook.Worksheets[newIndex];
                    }

                    // Set a meaningful worksheet name
                    sheet.Name = $"Sheet_{languages[i]}";

                    // Optionally set the workbook's culture (global setting)
                    workbook.Settings.CultureInfo = new CultureInfo(languages[i]);
                }

                // Save the workbook to a file
                string outputPath = "GlobalizationSettingsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
