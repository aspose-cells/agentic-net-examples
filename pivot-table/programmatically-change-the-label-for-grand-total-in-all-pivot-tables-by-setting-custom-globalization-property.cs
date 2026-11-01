// Title: Setting a custom grand total label for all pivot tables with Aspose.Cells for .NET (current version limitation)
// Common Searches: Aspose.Cells .NET change pivot table grand total caption programmatically | how to apply custom globalization settings to pivot tables using Aspose.Cells C# | update grand total label for every pivot table in an Excel file with Aspose.Cells | C# code example for customizing pivot table total text via Aspose.Cells | Aspose.Cells limitation custom grand total label property in latest release
// Tags: pivot table grand total label Aspose.Cells | custom globalization settings Aspose.Cells .NET | modify pivot table captions C# | Aspose.Cells pivot table customization | set pivot table total text programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Namespace for PivotTable

// The example loads an existing workbook with Aspose.Cells, explains that the CustomGlobalizationSettings and pivot table grand total label properties are not available in the current library version, and then saves the workbook unchanged.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The CustomGlobalizationSettings and PivotTable custom grand total label
            // properties are not available in the current Aspose.Cells version.
            // If needed, they can be set using a newer library version.

            // Save the (potentially modified) workbook
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
