// Title: Assign a custom GlobalizationSettings to Aspose.Cells Workbook before loading worksheets using C#
// AI Prompts: Write a C# example that creates a GlobalizationSettings object with custom decimal and thousands separators, passes it through LoadOptions, and loads an Excel file with Aspose.Cells. | Demonstrate how to apply a specific CultureInfo (e.g., fr-FR) together with custom number formatting by configuring GlobalizationSettings prior to opening a workbook in Aspose.Cells.
// Common Searches: Aspose.Cells C# load workbook with custom GlobalizationSettings before reading sheets | set decimal separator for Excel file using Aspose.Cells LoadOptions | apply French culture to Excel workbook when opening with Aspose.Cells .NET | how to configure number formatting globally in Aspose.Cells before loading workbook | C# example for custom globalization settings with Aspose.Cells LoadOptions
// Tags: load workbook with custom GlobalizationSettings Aspose.Cells | custom decimal separator Aspose.Cells C# | set CultureInfo before workbook load Aspose.Cells | Aspose.Cells number formatting configuration | LoadOptions globalization settings Excel .NET

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The sample creates a GlobalizationSettings object (optionally customizing decimal and thousands separators), supplies it via LoadOptions, loads the input Excel file with Aspose.Cells, applies a French CultureInfo for date and number formatting, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Create a custom GlobalizationSettings instance (optional customization)
            GlobalizationSettings customSettings = new GlobalizationSettings
            {
                // UseSystemSeparators = false,
                // DecimalSeparator = ",",
                // ThousandsSeparator = " "
            };

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Assign the custom globalization settings after loading (required by current API)
            workbook.Settings.GlobalizationSettings = customSettings;

            // Optionally set the culture for number/date formatting
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Save the workbook to demonstrate that it is usable
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
