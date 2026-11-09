// Title: Assign custom GlobalizationSettings (CultureInfo) to a Workbook before loading worksheets with Aspose.Cells for .NET
// Common Searches: how to set workbook globalization settings before loading worksheets in Aspose.Cells C# | Aspose.Cells assign French CultureInfo to workbook prior to opening file | initialize Aspose.Cells workbook with specific locale to control decimal separator | set Workbook.Settings.CultureInfo before reading Excel sheets using Aspose.Cells | apply custom globalization settings to Excel workbook in .NET with Aspose.Cells
// Tags: Workbook culture info configuration Aspose.Cells | regional French settings for workbook Aspose.Cells | custom globalization configuration for .NET Excel processing | load workbook using designated CultureInfo Aspose.Cells | globalizationsettings usage Aspose.Cells C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, assigns a French CultureInfo (fr-FR) to the workbook's globalization settings before any worksheets are accessed, and then saves the modified workbook, handling potential exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file.
                Workbook workbook = new Workbook(inputPath);

                // Apply custom globalization settings (French culture uses comma as decimal separator, etc.).
                workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

                // Save the workbook to the output file.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
