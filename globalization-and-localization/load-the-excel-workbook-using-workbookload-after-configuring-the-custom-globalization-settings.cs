// Title: Load an Excel workbook with British English (en-GB) CultureInfo using Aspose.Cells LoadOptions in C#
// AI Prompts: Generate C# code that creates a LoadOptions object with CultureInfo set to en-GB, loads an .xlsx file with Aspose.Cells, and saves the workbook. | Show how to switch the LoadOptions CultureInfo to another locale (e.g., fr-FR) while loading an Excel file with Aspose.Cells in .NET. | Provide a C# example that validates the input file path, applies custom globalization settings via LoadOptions, and handles exceptions during workbook loading.
// Common Searches: Aspose.Cells C# load workbook with specific CultureInfo en-GB | how to set locale for Excel import using LoadOptions in Aspose.Cells .NET | custom globalization settings for loading .xlsx files with Aspose.Cells | example of using LoadOptions.CultureInfo to load Excel in British English | error handling for missing Excel file when using Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions CultureInfo en-GB | C# load Excel workbook with locale | custom globalization settings for Excel import .NET | exception handling missing workbook Aspose.Cells | switch LoadOptions locale fr-FR Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

namespace AsposeCellsExample
{
    // The example checks for the presence of Input.xlsx, configures a LoadOptions object with LoadFormat.Xlsx and CultureInfo set to British English (en-GB), loads the workbook using these options, saves it as Output.xlsx, and includes exception handling for missing files or load errors.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Set custom culture (British English) for loading the workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    CultureInfo = new CultureInfo("en-GB")
                };

                // Load the workbook with the specified culture
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook to verify it was loaded correctly
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
